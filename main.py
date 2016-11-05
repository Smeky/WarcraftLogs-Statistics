import logger
import excel

import sys
import getopt
import urllib.request
import json
import xlsxwriter
import time

# -- Constants --

API_KEY = "api_key=e6b4169ddef610cca017c92d62d3028c"
URL_REPORT_FIGHTS   = "https://www.warcraftlogs.com/v1/report/fights/"
URL_REPORT_DMG_IN   = "https://www.warcraftlogs.com/v1/report/tables/damage-taken/"
URL_REPORT_BUFF     = "https://www.warcraftlogs.com/v1/report/tables/buffs/"
URL_REPORT_DEATH    = "https://www.warcraftlogs.com/v1/report/tables/deaths/"
URL_PARSES_CHAR     = "https://www.warcraftlogs.com/v1/parses/character/"
URL_ZONES           = "https://www.warcraftlogs.com/v1/zones"

CLASSES = ["DeathKnight", "DemonHunter", "Druid", "Hunter", "Mage", "Monk", "Paladin", "Priest", "Rogue", "Shaman", "Warlock", "Warrior"]

METRICS = ["dps", "hps"]

# Development data
DEV_VER = False      # If using development version (.txt data is used instead of links)
DEV_ZONES_FILE          = "zones.txt"
DEV_CLASSES_FILE        = "classes.txt"
DEV_REPORT_FIGHTS_FILE  = "report_fights.txt"
DEV_PARSE_CHAR_FILE     = "parse_char_smeky.txt"

# -- Globals --
# Parsing feature enabled/disabled
ENABLED_FEATURES = {
    "attended": True,
    "deaths":   True,
    "ilv":      True,
    "ranking":  True,
    "pots":     True,
    "flask":    True,
    "food":     True,
    "dmgin":    True,
    "dmgout":   True,
    "healin":   True,
    "healout":  True,
}

# Parsing fights enabled/disabled
ENABLED_STATISTIC = {
    "kills":    True,
    "Total":    True
}

REALMS = []

RAID_CODE       = "_unknown_"
RAID_DIFFICULTY = 5

player_data = {}    
fights_data = {}
raid_info   = {}
zone_info   = {}

logger = logger.Logger("logs/wls.txt")

# -- Functions --

def setFeatureState(feature, state):
    if feature in ENABLED_FEATURES:
        ENABLED_FEATURES[feature] = state
        return True
    
    return False

def setFeatureStateToAll(state):
    global ENABLED_FEATURES
    ENABLED_FEATURES = dict.fromkeys(ENABLED_FEATURES, state)


def printHelp():
    print("Help text not found")    # temp

def printVersion():
    print("Version text not found")    # temp    

def parseArgs():
    argv = sys.argv[1:]

    if len(argv) < 1:
        logger.log("Error: wrong amount of args was given. Most likely Raid code is missing")
        quit()

    global RAID_CODE
    RAID_CODE = argv[0]

    try:
        opts, args = getopt.getopt(argv[1:], "hvr:e:", ["parse_opts="])
    except getopt.GetoptError:
        logger.log("Error: arg exception")
        quit()

    for opt, arg in opts:
        if opt in ('-h', "-help"):
            printHelp()
        
        elif opt in ('-v', "-version"):
            printVersion() # Todo: add version

        elif opt in ('-r', "-realms"):
            to_parse = arg.split(',')

            for parse_opt in to_parse:
                REALMS.append(parse_opt)

        elif opt in ('-e', "-enable"):
            if arg == "all":
                setFeatureStateToAll(True)

                continue
            else:
                setFeatureStateToAll(False)

            if ',' not in arg:
                if setFeatureState(arg, True) == False:
                    logger.log("Error: Given unknown -enable arguemnt: " + arg)
                    quit()
                else:
                    continue

            to_parse = arg.split(',')

            for parse_opt in to_parse:
                # Try to set giben feature state
                # If this returns false, feature was not found
                if setFeatureState(parse_opt, True) == False:
                    logger.log("Error: Given unknown -enable option: " + parse_opt)
                    logger.log("  Skipping to next option")
                    continue




def buildUrl(base, *args):
    result  = base
    first   = True

    for arg in args:
        if first == True:
            result += '?'
        else:
            result += '&'

        first = False
        result += arg

    return result

def getJsonFromUrl(url, dev_file=""):
    if DEV_VER == True:
        if dev_file != "":
            with open(dev_file) as data_file:
                return json.load(data_file)


    try:
        response = urllib.request.urlopen(url)
        
        return json.loads(response.read().decode("utf-8"))

    except urllib.error.HTTPError as e:
        # Do not log missing json
        if e.code != 400:
            logger.log("Url respnse code: " + str(e.code))
            logger.log(e.reason)

        return False

def getCleanName(name):
    if '-' in name:
        return name.split('-', 1)
    else:
        return name

def setupPlayerList(friendlies):
    for character in friendlies:
        # Check if this character is a player
        if character.get('type', "Unknown") in CLASSES:
            data = {}

            data['name']    = character.get('name', "Unknown")
            data['class']   = character.get('type', "Unknown")
            data['id']      = character.get('id', 0)
            data['realm']   = REALMS[0] # Use first realm as deafult
            data['ranking'] = None
            data['fights']  = {}

            # Initialize data structure for all fights
            for fightID, _ in fights_data.items():
                data['fights'][fightID] = None

            for fight in character.get('fights', []):
                fightID    = fight['id']

                # Skip non-boss fights
                if fightID not in fights_data:
                    continue

                fight_data = {}

                fight_data['kill']          = fights_data[fightID]['kill']
                fight_data['difficulty']    = fights_data[fightID]['diff']
                fight_data['boss']          = fights_data[fightID]['boss']
                fight_data['dmg_taken']     = 0
                fight_data['dmg_done']      = 0
                fight_data['dps']           = 0
                fight_data['hps']           = 0
                fight_data['deaths']        = []

                fight_data['enhancements']  = {}
                fight_data['enhancements']['pot_1']     = False
                fight_data['enhancements']['pot_2']     = False
                fight_data['enhancements']['flask']     = 0
                fight_data['enhancements']['food']      = 0
                fight_data['enhancements']['food_type'] = ""

                data['fights'][fightID] = fight_data

            # Validate player - solves problem with logs that weren't ended properly
            for _, fight in data['fights'].items():
                if fight != None:
                    player_data[character['name']] = data
                    
                    break

# Store data from all fights
def parseFightsData():
    logger.log("\n-----------------------------")
    logger.log("Start: Parsing fights data")
    logger.log("-----------------------------")
    
    url = buildUrl(URL_REPORT_FIGHTS + RAID_CODE, API_KEY)

    logger.log("\nDownloading report for:\n" + url)
    report  = getJsonFromUrl(url, DEV_REPORT_FIGHTS_FILE)

    raid_info['title']      = report.get('title', "Unknown")
    raid_info['owner']      = report.get('owner', "Unknown")
    raid_info['raid_start'] = report.get('start', 0)
    raid_info['raid_end']   = report.get('end', 0)
    raid_info['zone']       = report.get('zone', 0)
    raid_info['code']       = RAID_CODE

    logger.log("\nRaid info:")
    logger.log("  title: " + raid_info['title'])
    logger.log("  start: " + str(raid_info['raid_start']))
    logger.log("  end: " + str(raid_info['raid_end']))

    logger.log("\nList of boss fights:")

    for fight in report['fights']:
        # Skip non-boss fights
        if fight['boss'] == 0:
            continue

        data = {}

        data['id']      = fight.get('id', 0)               # ID of the encounter
        data['start']   = fight.get('start_time', 0)       # Start time of the encounter
        data['end']     = fight.get('end_time', 0)         # End time of the encounter
        data['boss']    = fight.get('boss', 0)             # Boss WarcraftLogs ID 
        data['name']    = fight.get('name', "")            # Boss name
        data['kill']    = fight.get('kill', False)         # If boss was killed
        data['size']    = fight.get('size', 0)             # Raid size
        data['prcnt']   = fight.get('bossPercentage', 0)   # Boss percantage
        data['diff']    = fight.get('difficulty', 0)       # Raid size

        fights_data[data['id']] = data

        logger.log("  " + str(data['id']) + " " + data['name'])

    setupPlayerList(report['friendlies'])

    logger.log("\nList of attended players:")

    for name, _ in sorted(player_data.items()):
        logger.log("  " + name)

def parseZoneInfo():
    logger.log("\n-----------------------------")
    logger.log("Start: Parsing zone data")
    logger.log("-----------------------------")

    zone_ID = raid_info['zone']

    url = buildUrl(URL_ZONES, API_KEY)

    logger.log("Downloading zone data, zone ID: " + str(zone_ID))

    zone_data = getJsonFromUrl(url, DEV_ZONES_FILE)

    if zone_data == False:
        logger.log("Failed to download zone data from url:\n" + url)
        return False

    for raid in zone_data:
        if raid['id'] == zone_ID:
            zone_info['encounters'] = []
            zone_info['brackets']   = {}

            logger.log("\nEncounters:")

            boss_count = 1

            for encounter in raid['encounters']:
                encounterID = encounter['id']

                logger.log("  (" + str(encounterID) + ") " + encounter['name'])

                data = {}
                data['name']  = encounter['name']
                data['boss']  = encounterID
                data['order'] = boss_count

                zone_info['encounters'].append(data)

                boss_count += 1

            logger.log("\nBrackets:")

            for bracket in raid['brackets']:
                logger.log("  (" + str(bracket['id']) + ") " + bracket['name'])

                data = {}

                if '-' in bracket['name']:
                    bracket_ilv = bracket['name'].split('-')

                    data['min'] = int(bracket_ilv[0])
                    data['max'] = int(bracket_ilv[1])
                elif '+' in bracket['name']:
                    data['min'] = int(bracket['name'].split('+', 1)[0])
                    data['max'] = 99999

                zone_info['brackets'][bracket['id']] = data


def getBracketId(itemlevel):
    if itemlevel == 0:
        return 0

    for ID, bracket in zone_info['brackets'].items():
        if itemlevel >= bracket['min'] and itemlevel <= bracket['max']:
            return ID

    return 0

def getFightIdByBossId(bossID, kill=False):
    for ID, data in fights_data.items():
        if data['diff'] != RAID_DIFFICULTY:
            continue

        curr_boss_ID = data.get('boss', 0)

        if curr_boss_ID != 0:
            if kill == True:
                if curr_boss_ID == bossID and data.get('kill', False) == True:    
                    return ID
            else:
                if curr_boss_ID == bossID:
                    return ID

    return 0

def getBossIdByName(name):
    for encounter in zone_info['encounters']:
        if encounter['name'] == name:
            return encounter['boss']

    return 0

def getPlayerParseUrl(name, realm, metric):
    return buildUrl(URL_PARSES_CHAR + urllib.parse.quote(name) + '/' + realm + '/EU', metric, API_KEY)

def getPlayerBracketUrl(name, realm, metric, bracket):
    return buildUrl(URL_PARSES_CHAR + urllib.parse.quote(name) + '/' + realm + '/EU', metric, "bracket=" + str(bracket), API_KEY)

def validateParsedData(data):
    if data == False:
        return False

    for boss in data:
        for spec in boss['specs']:
            for spec_data in spec['data']:
                if validateStartTime(spec_data['start_time']) == True:
                    return True

    return False

def validateStartTime(start_time):
    min_time = raid_info['raid_start'] - 30000
    max_time = raid_info['raid_end'] + 30000

    return start_time >= min_time and start_time <= max_time

def getAllBrackets(char_data):
    result = None

    for boss in char_data:
        # Skip different difficulties
        if boss['difficulty'] != RAID_DIFFICULTY:
            continue

        for spec in boss['specs']:
            for data in spec['data']:
                if validateStartTime(data['start_time']) == True:
                    # Initialize dict
                    if result == None:
                        result = []

                    current_bracket = getBracketId(data['ilvl'])

                    # Do not add itemleves if it is duplicate
                    shouldAdd = True
                    for bracket in result:
                        if bracket == current_bracket:
                            shouldAdd = False
                            break

                    if shouldAdd == True:
                        result.append(current_bracket)


    return result

def parseBracketRanking(complete_ranking, bracket_data, metric):
    for boss in bracket_data:
        for spec in boss['specs']:
            if spec['combined'] != False:
                continue

            for data in spec['data']:
                # Check if the encounter was in this raid
                if validateStartTime(data['start_time']) != True:
                    continue

                boss_ranking = {}
                boss_ranking['prcnt'] = data['percent']
                boss_ranking['hist']  = data['historical_percent']
                boss_ranking['aps']   = data['persecondamount']

                bossID = getBossIdByName(boss['name'])

                if not bossID in complete_ranking:
                    complete_ranking[bossID] = {}

                if not 'ilv' in complete_ranking[bossID]:
                    complete_ranking[bossID]['ilv'] = data['ilvl']

                complete_ranking[bossID][metric] = boss_ranking

    return complete_ranking

def parseRankingForPlayers():
    if ENABLED_FEATURES['ranking'] == False:
        return

    logger.log("\n-----------------------------")
    logger.log("Start: Parsing ranking for all players")
    logger.log("-----------------------------\n")

    # Go through all players and check their rankings
    for name, player in sorted(player_data.items()):
        logger.log(name)
        
        ranking = None
        
        for metric in METRICS:
            log_msg = "  " + metric + ": "

            url         = getPlayerParseUrl(name, player['realm'], "metric=" + metric)
            char_data   = getJsonFromUrl(url)
            data_found  = validateParsedData(char_data)

            if data_found == False:
                for realm in REALMS:
                    url         = getPlayerParseUrl(name, realm, "metric=" + metric)
                    char_data   = getJsonFromUrl(url)

                    if validateParsedData(char_data) == True:
                        # Correct realm of player was found, store it
                        player_data[name]['realm'] = realm
                        data_found = True
                        break
                
            # If not data was found for the player in given realm list for this metric, skip him
            if data_found == False:
                log_msg += "No valid ranking found"
                logger.log(log_msg)

                continue

            # Get list of all itemleves so we can query for bracket ranking
            # All duplicates are removed
            brackets = getAllBrackets(char_data)

            if ranking == None:
                ranking = {}
            
            for bracket in brackets:
                url             = getPlayerBracketUrl(name, player['realm'], "metric=" + metric, bracket)
                ranking_data    = getJsonFromUrl(url)

                log_msg += str(bracket) + " "

                if ranking_data == False:
                    log_msg += " - Failed to retrieve bracket data"

                    continue
                else:
                    ranking = parseBracketRanking(ranking, ranking_data, metric)

            logger.log(log_msg)


        player['ranking'] = {}
        player['ranking'][RAID_DIFFICULTY] = ranking

        logger.log("-----------------------------")

def parsePotionsConsumable():
    POTIONS_ID = ["188027", "188028", "229206"]
    raid_start = raid_info['raid_start']
    raid_end   = raid_info['raid_end']

    for potion in POTIONS_ID:
        url_start   = "start="      + str(0)
        url_end     = "end="        + str(raid_end - raid_start)
        url_ability = "abilityid="  + potion
        
        url     = buildUrl(URL_REPORT_BUFF + RAID_CODE, url_start, url_end, url_ability, API_KEY)
        data    = getJsonFromUrl(url)

        total_buff_counter = 0

        for player_aura in data['auras']:
            name = player_aura['name']

            for buff in player_aura['bands']:
                fightID     = None
                start_time  = buff['startTime']

                for ID, fight_data in fights_data.items():
                    if start_time >= fight_data['start'] and start_time < fight_data['end']:
                        fightID = ID
                        break

                # Skip non-boss fights
                if fightID == None or player_data[name]['fights'][fightID] == None:
                    continue

                enhancements = player_data[name]['fights'][fightID]['enhancements']

                # If this was first potion used in this fight
                if enhancements['pot_1'] == False:
                    enhancements['pot_1'] = True
                else:
                    enhancements['pot_2'] = True

                player_data[name]['fights'][fightID]['enhancements'] = enhancements

                total_buff_counter += 1

        logger.log("  Potion: " + potion + " - Total of " + str(total_buff_counter) + " buffs found")

def parseFlaskConsumable():
    FLASK_ID = ["188031", "188033", "188035"]
    raid_start = raid_info['raid_start']
    raid_end   = raid_info['raid_end']

    for flask in FLASK_ID:
        url_start   = "start="      + str(0)
        url_end     = "end="        + str(raid_end - raid_start)
        url_ability = "abilityid="  + flask
        
        url     = buildUrl(URL_REPORT_BUFF + RAID_CODE, url_start, url_end, url_ability, API_KEY)
        data    = getJsonFromUrl(url)

        total_buff_counter = 0

        for player_aura in data['auras']:
            name = player_aura['name']

            for buff in player_aura['bands']:
                fightID     = None
                start_time  = buff['startTime']
                end_time    = buff['endTime']
                uptime      = 0

                for ID, fight_data in fights_data.items():
                    if start_time >= fight_data['start'] and end_time <= fight_data['end']:
                        uptime = (end_time - start_time) / (fight_data['end'] - fight_data['start'])

                        fightID = ID
                        break

                # Skip non-boss fights
                if fightID == None or player_data[name]['fights'][fightID] == None:
                    continue

                player_data[name]['fights'][fightID]['enhancements']['flask'] = uptime

                total_buff_counter += 1

        logger.log("  Flask: " + flask + " - Total of " + str(total_buff_counter) + " buffs found")

def parseFoodConsumable():
    FOOD_ID = {
        "300": ["225597", "225598", "225599"],
        "375": ["225602", "225603", "225604", "225605", "225606"] 
    }

    raid_start = raid_info['raid_start']
    raid_end   = raid_info['raid_end']

    for category, food_ids in FOOD_ID.items():
        for food in food_ids:
            url_start   = "start="      + str(0)
            url_end     = "end="        + str(raid_end - raid_start)
            url_ability = "abilityid="  + food
            
            url     = buildUrl(URL_REPORT_BUFF + RAID_CODE, url_start, url_end, url_ability, API_KEY)
            data    = getJsonFromUrl(url)

            total_buff_counter = 0

            for player_aura in data['auras']:
                name = player_aura['name']

                for buff in player_aura['bands']:
                    fightID     = None
                    start_time  = buff['startTime']
                    end_time    = buff['endTime']
                    uptime      = 0

                    for ID, fight_data in fights_data.items():
                        if start_time >= fight_data['start'] and end_time <= fight_data['end']:
                            uptime = (end_time - start_time) / (fight_data['end'] - fight_data['start'])

                            fightID = ID
                            break

                    # Skip non-boss fights
                    if fightID == None or player_data[name]['fights'][fightID] == None:
                        continue

                    player_data[name]['fights'][fightID]['enhancements']['food']        = uptime
                    player_data[name]['fights'][fightID]['enhancements']['food_type']   = category

                    total_buff_counter += 1

            logger.log("  Food: " + food + " - Total of " + str(total_buff_counter) + " buffs found")

def parseConsumablesInfo():
    if ENABLED_FEATURES['pots'] == False and ENABLED_FEATURES['flask'] == False and ENABLED_FEATURES['food'] == False:
        return

    logger.log("\n-----------------------------")
    logger.log("Start: Parsing consumables information for all players")
    logger.log("-----------------------------\n")

    # https://www.warcraftlogs.com:443/v1/report/tables/buffs/49twcxraQDzNXB2R?start=0&end=9999999&abilityid=188027&api_key=e6b4169ddef610cca017c92d62d3028c

    if ENABLED_FEATURES['pots'] == True:
        parsePotionsConsumable()
    if ENABLED_FEATURES['flask'] == True:
        parseFlaskConsumable()
    if ENABLED_FEATURES['food'] == True:
        parseFoodConsumable()

def parseDeathInfo():
    if ENABLED_FEATURES['deaths'] == False:
        return

    logger.log("\n-----------------------------")
    logger.log("Start: Parsing death statistics")
    logger.log("-----------------------------\n")

    raid_start = raid_info['raid_start']
    raid_end   = raid_info['raid_end']

    url_start   = "start="      + str(0)
    url_end     = "end="        + str(raid_end - raid_start)
    
    url     = buildUrl(URL_REPORT_DEATH + RAID_CODE, url_start, url_end, API_KEY)
    data    = getJsonFromUrl(url)

    total_death_count = 0
    for entry in data['entries']:
        name    = entry['name']
        fightID = entry['fight']

        # Skip non-boss fights
        if fightID not in player_data[name]['fights']:
            continue

        death_data = {}

        # For now, all we need is to increase length of the list so we can track that player died
        # Todo: Add some death info like abilities, healing, reductions, etc.

        player_data[name]['fights'][fightID]['deaths'].append(death_data)

        total_death_count += 1

    logger.log("total of " + str(total_death_count) + " deaths found")

def parseDamageTaken():
    if ENABLED_FEATURES['dmgin'] == False:
        return

    logger.log("\n-----------------------------")
    logger.log("Start: Parsing damage taken")
    logger.log("-----------------------------\n")

# -- Program start --

parseArgs()

logger.log("-------------------------------------------")
logger.log("Beginning analysis for [ " + RAID_CODE + " ]")
logger.log("-------------------------------------------")

logger.log("Realms:")
for realm in REALMS:
    logger.log("  " + realm)

logger.log("\nFeatures:")
for feature, enabled in ENABLED_FEATURES.items():
    if enabled == True:
        logger.log("  " + feature)

parseFightsData()
parseZoneInfo()
parseRankingForPlayers()
parseConsumablesInfo()
parseDeathInfo()
parseDamageTaken()

logger.log("\n-----------------------------")
logger.log("Writing parsed data into .xlsx file")
logger.log("-----------------------------")

excel = excel.ExcelTable(zone_info, raid_info, fights_data, player_data)

timestamp = time.strftime("%Y%m%d-%H%M%S")
filename  = "result/" + timestamp + ".xlsx"

excel.openFile(filename)

offset_x    = 2
offset_y    = 2

excel.writeStatisticsTable(offset_x + 1, offset_y)
excel.writeRankingTable(RAID_DIFFICULTY, offset_x, offset_y)
excel.writeEncounterStats(RAID_DIFFICULTY, offset_x, offset_y)

excel.closeFile()

logger.log('\n-------------------------------------------')
logger.log('Analysis ended successfully')
logger.log('-------------------------------------------')