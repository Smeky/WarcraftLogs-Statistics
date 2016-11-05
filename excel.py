import logger 
import xlsxwriter

CLASS_COLOR = {}
CLASS_COLOR['DeathKnight']  = '#C41F3B'
CLASS_COLOR['DemonHunter']  = '#A330C9'
CLASS_COLOR['Druid']        = '#FF7D0A'
CLASS_COLOR['Hunter']       = '#ABD473'
CLASS_COLOR['Mage']         = '#69CCF0'
CLASS_COLOR['Monk']         = '#00FF96'
CLASS_COLOR['Paladin']      = '#F58CBA'
CLASS_COLOR['Priest']       = '#FFFFFF'
CLASS_COLOR['Rogue']        = '#FFF569'
CLASS_COLOR['Shaman']       = '#0070DE'
CLASS_COLOR['Warlock']      = '#9482C9'
CLASS_COLOR['Warrior']      = '#C79C6E'

METRIC_TYPE = ['dps', 'hps']

METRIC_COLOR = {}
METRIC_COLOR['dps']         = '#366092'
METRIC_COLOR['hps']         = '#7DB545'

_WS_STATISTICS = 0          # Statistics worksheet index
_WS_RANKING = 1             # Ranking worksheet index
_WS_BOSS_INDEX_INIT = 2     # At which number should boss indexing begin
_WS_INDEX_BY_BOSS = {}      # List of worksheet indices stored under Boss ID

class ExcelTable:
    _worksheet      = []
    _player_count   = {}

    def __init__(self, zone_info, raid_info, fights_data, player_data):
        # Store all required data
        self._zone_info   = zone_info
        self._raid_info   = raid_info
        self._fights_data = fights_data
        self._player_data = player_data

        # Dimension additions
        self._width_additions  = 3  # Names, Total, Avg. ILv
        self._height_additions = 3  # Title, padding

        self._player_count['dps'] = 0
        self._player_count['hps'] = 0

        # Init logger
        self._logger = logger.Logger("logs/wls_excel.txt")

    # Add new worksheet to the workbook and return its index
    def _addWorksheet(self, name=""):
        log_msg = ""

        if name != "":
            log_msg += "Adding '" + name + "' worksheet"
            self._worksheet.append(self._workbook.add_worksheet(name))
        else:
            log_msg += "Adding unnamed worksheet"
            self._worksheet.append(self._workbook.add_worksheet())

        index = len(self._worksheet) - 1 
        
        self._logger.log(log_msg + "(idx " + str(index) + ")") 

        return index

    def getRankingTableSize(self, metric):
        width   = self._width_additions  + len(self._zone_info['encounters'])
        height  = 0

        if metric == "dps":
            height = self._height_additions + self._player_count['dps']
        elif metric == "hps":
            height = self._height_additions + self._player_count['hps']

        # Todo: Calculate correct dimensions
        return {'w': width, 'h': height}

    def openFile(self, name):
        self._logger.log("\nOpening file: " + name)

        self._workbook = xlsxwriter.Workbook(name)

        self._logger.log("-----------------------------")

        # Add Statistic worksheet
        _WS_STATISTICS = self._addWorksheet("Statistics")
        self._worksheet[_WS_STATISTICS].set_tab_color('#366092')

        # Add Ranking worksheet
        _WS_RANKING = self._addWorksheet("Ranking")
        self._worksheet[_WS_RANKING].set_tab_color('#366092')

        # Add worksheet for every boss
        for encounter in self._zone_info['encounters']:
            index = self._addWorksheet(self._shrinkBossName(encounter['name']))

            fight_attended = False
            for fightID, fight_data in self._fights_data.items():
                if fight_data['boss'] != encounter['boss']:
                    continue
                else:
                    fight_attended = True

                if fight_data['kill'] != True:
                    self._worksheet[index].set_tab_color('#EB0000')
                else:
                    self._worksheet[index].set_tab_color('#92D050')

            if fight_attended == False:
                self._worksheet[index].set_tab_color('#EB0000')


            _WS_INDEX_BY_BOSS[encounter['boss']] = index


    def closeFile(self):
        self._logger.log("\n-----------------------------")
        self._logger.log("Closing file")
        self._logger.log("-----------------------------")

        self._workbook.close()

    def _getKillCount(self):
        counter = 0

        for _, fight in self._fights_data.items():
            if fight['kill'] == True:
                counter += 1

        return counter

    def _getPlayerAverageItemlevel(self, player, difficulty=None):
        if player['ranking'] == None:
            return 0

        if difficulty != None and difficulty not in player['ranking']:
            return 0

        ilvs = []

        if difficulty != None:
            for bossID, boss_data in player['ranking'][difficulty].items():
                ilvs.append(boss_data['ilv'])
        else:
            for _, diff_data in player['ranking'].items():
                if diff_data == None:
                    continue

                for bossID, boss_data in diff_data.items():
                    ilvs.append(boss_data['ilv'])

        if len(ilvs) == 0 or sum(ilvs) == 0:
            return 0
        else:
            return round(sum(ilvs) / len(ilvs), 1)

    def _getPlayerAttendedFights(self, player, kill=False):
        counter = 0

        for _, fight in player['fights'].items():
            if fight == None:
                continue

            if kill == True:
                if fight['kill'] == True:
                    counter += 1
            else:
                counter += 1

        return counter

    def _getPlayerAverageRanking(self, player, metric):
        # Todo: handle player's role
        #       -> Players found in healing ranking shouldn't be considered as DPS
        #       -> This will have problem with offspecs

        total   = 0
        counter = 0

        # Check if player has some valid ranking
        if player['ranking'] == None:
            return None

        for _, diff in player['ranking'].items():
            if diff == None:
                continue

            for _, boss in diff.items():
                if metric in boss:
                    total += boss[metric]['hist']
                    counter += 1

        # Return average
        if total != 0:
            return total / counter
        
        return None

    def _getPlayerDeaths(self, player, kill=False):
        counter = 0

        for _, fight in player['fights'].items():
            if fight == None:
                continue

            if kill == True:
                if fight['kill'] == True:
                    counter += len(fight['deaths'])
            else:
                counter += len(fight['deaths'])

        return counter

    def _getPlayerTotalPotionUsed(self, player, kill=False):
        counter = 0

        for _, fight in player['fights'].items():
            if fight == None:
                continue

            if kill == True:
                if fight['kill'] == True:
                    if fight['enhancements']['pot_1'] == True:
                        counter += 1

                    if fight['enhancements']['pot_2'] == True:
                        counter += 1
            else:
                if fight['enhancements']['pot_1'] == True:
                    counter += 1

                if fight['enhancements']['pot_2'] == True:
                    counter += 1

        return counter        

    def _getPlayerTotalDamageTaken(self, player, kill=False):
        result = 0

        for _, fight in player['fights'].items():
            if fight == None:
                continue

            if kill == True:
                if fight['kill'] == True:
                    result += fight['dmg_taken']

            else:
                result += fight['dmg_taken']

        return result

    def _getPlayerTotalFlaskUptime(self, player, kill=False):
        uptime = 0
        counter = 0

        for _, fight in player['fights'].items():
            if fight == None:
                continue

            if kill == True:
                if fight['kill'] == True:
                    uptime += fight['enhancements']['flask']                        
                    counter += 1
            else:
                uptime += fight['enhancements']['flask']
                counter += 1

        if uptime != 0 and counter != 0:
            return (uptime / counter) * 100
        else:
            return 0

    def _getPlayerTotalFoodUptime(self, player, kill=False):
        uptime = 0
        counter = 0

        for _, fight in player['fights'].items():
            if fight == None:
                continue

            if kill == True:
                if fight['kill'] == True:
                    uptime += fight['enhancements']['food']
                    counter += 1
            else:
                uptime += fight['enhancements']['food']
                counter += 1

        if uptime != 0 and counter != 0:
            return (uptime / counter) * 100
        else:
            return 0

    def _shrinkBossName(self, full_name):
        HARD_CAP = 9
        SOFT_CAP = 5

        char_index = None

        if ' ' in full_name:
            char_index = full_name.index(' ')
        elif ',' in full_name:
            char_index = full_name.index(',')

        if char_index != None:
            # Apply softcap 
            if char_index < HARD_CAP and char_index >= SOFT_CAP:
                return full_name[:char_index] + "..."
            # Apply hardcap
            else:
                return full_name[:HARD_CAP] + "..."
        else:
            # Apply hardcap
            if len(full_name) > HARD_CAP:
                return full_name[:HARD_CAP] + "..."
            else:
                return full_name

    def _writeRankingTableTitle(self, metric, offset_x, offset_y):
        # Ranking title
        title_format = self._workbook.add_format({
            'bold':         1, 
            'border':       1,
            'align':        'center',
            'valign':       'vcenter',
            'font_size':    16,
            'font_color':   'white',
            'bg_color':     METRIC_COLOR.get(metric, '#FFFFFF')
        })

        dimensions = self.getRankingTableSize(metric)
        self._worksheet[_WS_RANKING].merge_range(offset_y, offset_x + 1, offset_y + 1, offset_x + dimensions['w'] - 1, metric.upper() + " (Ranking)", title_format)

    def _writeBossNameRow(self, offset_x, offset_y):
        boss_name_format = self._workbook.add_format({
            'align':        "center", 
            'valign':       "vcenter", 
            'bold':         1, 
            'font_size':    12, 
            'border':       1, 
            'font_color':   'white', 
            'bg_color':     "#244062"})

        rank_total_format = self._workbook.add_format({
            'align':        "center", 
            'valign':       "vcenter", 
            'bold':         1, 
            'font_size':    12, 
            'border':       1, 
            'font_color':   'white', 
            'bg_color':     "#C0504D"})

        avrg_ilv_format = self._workbook.add_format({
            'align':        "center", 
            'valign':       "vcenter", 
            'bold':         1, 
            'font_size':    12, 
            'border':       1, 
            'font_color':   'white', 
            'bg_color':     "#366092"})

        boss_row = offset_y + 3
        boss_col = offset_x + 1
        boss_idx = 0
        boss_count = len(self._zone_info['encounters'])

        self._worksheet[_WS_RANKING].set_row(boss_row, 18)
        self._worksheet[_WS_RANKING].set_column(offset_x + 0,              offset_x + 0,              15)  # Player names
        self._worksheet[_WS_RANKING].set_column(offset_x + 1,              offset_x + boss_count,     14)  # Boss names
        self._worksheet[_WS_RANKING].set_column(offset_x + boss_count + 1, offset_x + boss_count + 1, 13)  # Total
        self._worksheet[_WS_RANKING].set_column(offset_x + boss_count + 2, offset_x + boss_count + 2, 13)  # Avg. ILv

        # Todo: Add automatic name shortening
        for encounter in self._zone_info['encounters']:
            self._worksheet[_WS_RANKING].write(boss_row, boss_col, self._shrinkBossName(encounter['name']), boss_name_format)
            boss_col += 1

        self._worksheet[_WS_RANKING].write(boss_row, boss_col,     "Total",    rank_total_format)
        self._worksheet[_WS_RANKING].write(boss_row, boss_col + 1, "Avg. ILv", avrg_ilv_format)

    def _hasMetricRanking(self, player, metric, difficulty):
        ranking = player['ranking'][difficulty]
                
        for boss in self._zone_info['encounters']:
            # Skip kills that player didn't attend
            if boss['boss'] in ranking:
                boss_ranking = ranking[boss['boss']]

                if metric in boss_ranking:
                    return True

        return False

    def _formatRankValue(self, value):
        if value == None:
            return '-'
        else:
            return round(value, 2)

    def _writeStatisticsInfo(self, offset_x, offset_y):
        cell_rinfo_head      = self._workbook.add_format({'border': 1, 'bold': 1, 'font_size': 14, 'font_color': 'white', 'bg_color': "#2F75B5"})
        cell_rinfo_left      = self._workbook.add_format({'border': 1, 'bg_color': "#F2F2F2"})
        cell_rinfo_right     = self._workbook.add_format({'right': 1})
        cell_rinfo_right_bot = self._workbook.add_format({'right': 1, 'bottom': 1})

        raid_start  = self._raid_info['raid_start']
        raid_end    = self._raid_info['raid_end']
        raid_length = raid_end - raid_start
        raid_link   = "https://www.warcraftlogs.com/reports/" + self._raid_info['code']

        # Raid info header
        self._worksheet[_WS_STATISTICS].merge_range(offset_y, offset_x, offset_y, offset_x + 2, "Development info", cell_rinfo_head)
        # Uploader
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x, "Raid title", cell_rinfo_left)
        self._worksheet[_WS_STATISTICS].merge_range(offset_y + 1, offset_x + 1, offset_y + 1, offset_x + 2, self._raid_info['title'], cell_rinfo_right)
        # Uploader
        self._worksheet[_WS_STATISTICS].write(offset_y + 2, offset_x, "Uploader", cell_rinfo_left)
        self._worksheet[_WS_STATISTICS].merge_range(offset_y + 2, offset_x + 1, offset_y + 2, offset_x + 2, self._raid_info['owner'], cell_rinfo_right)
        # Time length
        self._worksheet[_WS_STATISTICS].write(offset_y + 3, offset_x, "Time length", cell_rinfo_left)
        self._worksheet[_WS_STATISTICS].merge_range(offset_y + 3, offset_x + 1, offset_y + 3, offset_x + 2, raid_length, cell_rinfo_right)
        # Start / End
        self._worksheet[_WS_STATISTICS].write(offset_y + 4, offset_x, "Start / End", cell_rinfo_left)
        self._worksheet[_WS_STATISTICS].merge_range(offset_y + 4, offset_x + 1, offset_y + 4, offset_x + 2, str(raid_start) + " / " + str(raid_end), cell_rinfo_right)
        # Raid code
        self._worksheet[_WS_STATISTICS].write(offset_y + 5, offset_x, "Raid code", cell_rinfo_left)
        self._worksheet[_WS_STATISTICS].merge_range(offset_y + 5, offset_x + 1, offset_y + 5, offset_x + 2, self._raid_info['code'], cell_rinfo_right)
        # Link (warcraftlogs)
        self._worksheet[_WS_STATISTICS].write(offset_y + 6, offset_x, "Link", cell_rinfo_left)
        self._worksheet[_WS_STATISTICS].merge_range(offset_y + 6, offset_x + 1, offset_y + 6, offset_x + 2, raid_link, cell_rinfo_right_bot)

    def _writeStatisticsHeaders(self, offset_x, offset_y):
        offset_x += 1

        # Set category and data type row height
        self._worksheet[_WS_STATISTICS].set_row(offset_y + 0, 18)
        self._worksheet[_WS_STATISTICS].set_row(offset_y + 1, 18)

        # General
        cat_general_format = self._workbook.add_format({'align': "center", 'valign': "vcenter", 'bold': 1, 'border': 1, 'font_color': 'white', 'bg_color': "#16365C", 'font_size': 14})
        col_general_format = self._workbook.add_format({'align': "center", 'valign': "vcenter", 'bold': 1, 'border': 1, 'font_color': 'white', 'bg_color': "#16365C"})

        self._worksheet[_WS_STATISTICS].merge_range(offset_y, offset_x, offset_y, offset_x + 2, "General", cat_general_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 0, "Attended", col_general_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 1, "Deaths", col_general_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 2, "Avg. iLv", col_general_format)

        offset_x += 3

        # Consumables
        cat_consumables_format = self._workbook.add_format({'align': "center", 'valign': "vcenter", 'bold': 1, 'border': 1, 'font_color': 'white', 'bg_color': "#3A7A3C", 'font_size': 14})
        col_consumables_format = self._workbook.add_format({'align': "center", 'valign': "vcenter", 'bold': 1, 'border': 1, 'font_color': 'white', 'bg_color': "#3A7A3C"})

        self._worksheet[_WS_STATISTICS].merge_range(offset_y, offset_x, offset_y, offset_x + 2, "Consumables", cat_consumables_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 0, "Potions used", col_consumables_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 1, "Flask uptime", col_consumables_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 2, "Food uptime", col_consumables_format)

        offset_x += 3

        # Average ranking
        cat_avgrank_format = self._workbook.add_format({'align': "center", 'valign': "vcenter", 'bold': 1, 'border': 1, 'font_color': 'white', 'bg_color': "#297283", 'font_size': 14})
        col_avgrank_format = self._workbook.add_format({'align': "center", 'valign': "vcenter", 'bold': 1, 'border': 1, 'font_color': 'white', 'bg_color': "#297283"})

        self._worksheet[_WS_STATISTICS].merge_range(offset_y, offset_x, offset_y, offset_x + 2, "Average ranking", cat_avgrank_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 0, "DPS", col_avgrank_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 1, "HSP", col_avgrank_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 2, "Surviv", col_avgrank_format)

        offset_x += 3

        # In / out
        cat_inout_format = self._workbook.add_format({'align': "center", 'valign': "vcenter", 'bold': 1, 'border': 1, 'font_color': 'white', 'bg_color': "#963634", 'font_size': 14})
        col_inout_format = self._workbook.add_format({'align': "center", 'valign': "vcenter", 'bold': 1, 'border': 1, 'font_color': 'white', 'bg_color': "#963634"})

        self._worksheet[_WS_STATISTICS].merge_range(offset_y, offset_x, offset_y, offset_x + 2, "In / out", cat_inout_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 0, "Dmg Done", col_inout_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 1, "Heal Done", col_inout_format)
        self._worksheet[_WS_STATISTICS].write(offset_y + 1, offset_x + 2, "Dmg Taken", col_inout_format)

    def _writeStatisticsData(self, offset_x, offset_y, kill_data):        
        # Write category headers
        self._writeStatisticsHeaders(offset_x, offset_y)
        offset_y += 2 # 2x header row

        player_row   = offset_y
        player_count = 1
        for name, player in sorted(self._player_data.items()):
            # Player name format
            name_bg_color   = "#323232" if (player_count % 2 == 0) else "#262626"
            name_font_color = CLASS_COLOR[player['class']]

            cell_name_format = self._workbook.add_format({'bold': 1, 'bg_color': name_bg_color, 'font_color': name_font_color})

            # Player Name
            self._worksheet[_WS_STATISTICS].write(player_row, offset_x, name, cell_name_format)

            # -- Statistics --
            stat_column = offset_x + 1

            # General
            stat_bg_color = "#D1DFEF" if (player_count % 2 == 0) else "#9DB9DB"
        
            cell_general_left_format    = self._workbook.add_format({'align': 'center', 'bg_color': stat_bg_color, 'right': 1, 'right_color': "#808080"})
            cell_general_mid_format     = self._workbook.add_format({'align': 'center', 'bg_color': stat_bg_color, 'right': 1, 'right_color': "#808080"})
            cell_general_right_format   = self._workbook.add_format({'indent': 1, 'bg_color': stat_bg_color, 'right': 1})

            attended    = self._getPlayerAttendedFights(player, kill=kill_data)
            deaths      = self._getPlayerDeaths(player, kill=kill_data)
            avg_ilv     = self._getPlayerAverageItemlevel(player)

            if avg_ilv == 0:
                avg_ilv = '-'
                cell_general_right_format.set_indent(0)
                cell_general_right_format.set_align('center')
            
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 0, attended, cell_general_left_format)
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 1, deaths, cell_general_mid_format)
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 2, avg_ilv, cell_general_right_format)

            # Consumables
            stat_column += 3
            stat_bg_color = "#D5EBD6" if (player_count % 2 == 0) else "#AAD6AB"
            
            cell_consumables_left_format    = self._workbook.add_format({'align': 'center', 'bg_color': stat_bg_color, 'right': 1, 'right_color': "#808080"})
            cell_consumables_mid_format     = self._workbook.add_format({'align': 'center', 'bg_color': stat_bg_color, 'right': 1, 'right_color': "#808080"})
            cell_consumables_right_format   = self._workbook.add_format({'align': 'center', 'bg_color': stat_bg_color, 'right': 1})

            potions         = self._getPlayerTotalPotionUsed(player, kill=kill_data)
            flask_uptime    = str(round(self._getPlayerTotalFlaskUptime(player, kill=kill_data), 0)) + '%'
            food_uptime     = str(round(self._getPlayerTotalFoodUptime(player, kill=kill_data), 0)) + '%'
            
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 0, potions, cell_consumables_left_format)
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 1, flask_uptime, cell_consumables_mid_format)
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 2, food_uptime, cell_consumables_right_format)

            # Average ranking
            stat_column += 3
            stat_bg_color = "#D9EEF3" if (player_count % 2 == 0) else "#ADDAE5"

            cell_avgrank_left_format    = self._workbook.add_format({'indent': 1, 'bg_color': stat_bg_color, 'right': 1, 'right_color': "#808080"})
            cell_avgrank_mid_format     = self._workbook.add_format({'indent': 1, 'bg_color': stat_bg_color, 'right': 1, 'right_color': "#808080"})
            cell_avgrank_right_format   = self._workbook.add_format({'indent': 1, 'bg_color': stat_bg_color, 'right': 1})

            avg_rank_dps    = self._formatRankValue(self._getPlayerAverageRanking(player, "dps"))
            avg_rank_hps    = self._formatRankValue(self._getPlayerAverageRanking(player, "hps"))
            avg_rank_surviv = self._formatRankValue(self._getPlayerAverageRanking(player, "surviv"))

            if avg_rank_dps == '-':
                cell_avgrank_left_format.set_indent(0)
                cell_avgrank_left_format.set_align('center')

            if avg_rank_hps == '-':
                cell_avgrank_mid_format.set_indent(0)
                cell_avgrank_mid_format.set_align('center')

            if avg_rank_surviv == '-':
                cell_avgrank_right_format.set_indent(0)
                cell_avgrank_right_format.set_align('center')
            
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 0, avg_rank_dps, cell_avgrank_left_format)
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 1, avg_rank_hps, cell_avgrank_mid_format)
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 2, avg_rank_surviv, cell_avgrank_right_format)

            # In / out
            stat_column += 3
            stat_bg_color = "#FFD9D9" if (player_count % 2 == 0) else "#FFAFAF"

            cell_inout_left_format    = self._workbook.add_format({'align': 'center', 'bg_color': stat_bg_color, 'right': 1, 'right_color': "#808080"})
            cell_inout_mid_format     = self._workbook.add_format({'align': 'center', 'bg_color': stat_bg_color, 'right': 1, 'right_color': "#808080"})
            cell_inout_right_format   = self._workbook.add_format({'align': 'center', 'bg_color': stat_bg_color, 'right': 1})

            dmg_done    = '-'
            heal_done   = '-'
            dmg_taken   = '-'
            
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 0, dmg_done, cell_inout_left_format)
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 1, heal_done, cell_inout_mid_format)
            self._worksheet[_WS_STATISTICS].write(player_row, stat_column + 2, dmg_taken, cell_inout_right_format)

            player_row += 1
            player_count += 1

        # end forloop

        return player_row + 2 # return at which row we have ended 

    def writeStatisticsTable(self, main_offset_x, main_offset_y):
        self._logger.log("\nCreating general statistics table")
        self._logger.log("-----------------------------")

        offset_x = main_offset_x
        offset_y = main_offset_y

        # Set width for all columns
        self._worksheet[_WS_STATISTICS].set_row(offset_y, 18)
        self._worksheet[_WS_STATISTICS].set_column(offset_x + 0, offset_x + 0, 15)  # Player names
        self._worksheet[_WS_STATISTICS].set_column(offset_x + 1, offset_x + 12, 14)  # Attended fights
        
        kill_count = self._getKillCount()

        if kill_count > 0:
            # Write header for 'KILLS' statistics
            format_props = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 16, 'bold': 1, 'bg_color': "#D4ECBA"}
            cell_title_kills = self._workbook.add_format(format_props)

            self._worksheet[_WS_STATISTICS].merge_range(offset_y, offset_x, offset_y + 1, offset_x + 2, "Kills (" + str(kill_count) + ")", cell_title_kills)
            offset_y += 3 # column height + offset

            offset_y = self._writeStatisticsData(offset_x, offset_y, True)


        # Write header for 'TOTAL' statistics
        format_props = {'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 16, 'bold': 1, 'bg_color': "#FCD5B4"}
        cell_title_total = self._workbook.add_format(format_props)

        self._worksheet[_WS_STATISTICS].merge_range(offset_y, offset_x, offset_y + 1, offset_x + 2, "Total (" + str(len(self._fights_data)) + ")", cell_title_total)
        offset_y += 3 # column height + offset

        offset_y = self._writeStatisticsData(offset_x, offset_y, False)


        self._writeStatisticsInfo(offset_x, offset_y + 3)


    def writeRankingTable(self, difficulty, main_offset_x, main_offset_y):
        offset_x = main_offset_x
        offset_y = main_offset_y

        self._logger.log("\nCreating ranking tables")
        self._logger.log("-----------------------------")

        for metric in METRIC_TYPE:
            self._logger.log("  " + metric + " table at " + str(offset_x) + "x" + str(offset_y))

            self._writeRankingTableTitle(metric, offset_x, offset_y)
            self._writeBossNameRow(offset_x, offset_y)

            cell_total_format  = self._workbook.add_format({'border': 1, 'bold': 1, 'align': 'center', 'bg_color': "#E6B8B7"})
            cell_avgilv_format = self._workbook.add_format({'border': 1, 'bold': 1, 'align': 'left', 'indent': 1, 'bg_color': "#B8CCE4"})
            cell_rank_format_odd  = self._workbook.add_format({'align': 'center', 'bg_color': "#D9D9D9"})
            cell_rank_format_even = self._workbook.add_format({'align': 'center', 'bg_color': "#FFFFFF"})

            player_row   = offset_y + 4
            player_count = 0

            for name, player in sorted(self._player_data.items()):
                # Skip players with no ranking
                if player['ranking'] == None:
                    continue

                # If player hasn't attended any kills on current difficulty 
                if player['ranking'].get(difficulty, None) == None:
                    continue

                # Skip players that do not have any ranking in this metric
                if self._hasMetricRanking(player, metric, difficulty) == False:
                    continue

                name_bg_color = "#323232" if (player_row % 2 == 0) else "#262626"

                cell_name_format = self._workbook.add_format({'bold': 1, 'bg_color': "#323232", 'font_color': CLASS_COLOR[player['class']], 'bg_color': name_bg_color})
                cell_rank_format = cell_rank_format_even if (player_row % 2 == 0) else cell_rank_format_odd

                boss_column = offset_x + 1

                # Todo: Write name here
                self._worksheet[_WS_RANKING].write(player_row, offset_x, name, cell_name_format)

                ranking = player['ranking'][difficulty]
                
                # Write ranking for every boss
                for boss in self._zone_info['encounters']:
                    rank = '-'

                    # Skip kills that player didn't attend
                    if boss['boss'] in ranking:
                        boss_ranking = ranking[boss['boss']]

                        if metric in boss_ranking:
                            rank = boss_ranking[metric]['hist']

                    self._worksheet[_WS_RANKING].write(player_row, boss_column, rank, cell_rank_format)

                    boss_column += 1

                # Total Formula
                form_range = xlsxwriter.utility.xl_range(player_row, offset_x + 1, player_row, boss_column - 1)
                self._worksheet[_WS_RANKING].write_formula(player_row, boss_column, "=ROUND(AVERAGE(" + form_range + "),2)", cell_total_format)
                
                # Average Itemlevel
                self._worksheet[_WS_RANKING].write(player_row, boss_column + 1, self._getPlayerAverageItemlevel(player, difficulty), cell_avgilv_format)

                player_row += 1
                player_count += 1

            # Store highest player count for table height
            if player_count > self._player_count[metric]:
                self._player_count[metric] = player_count

            offset_y += self.getRankingTableSize(metric)['h'] + 4

    def _writeBossStatsRow(self, sheet_index, offset_x, offset_y):
        default_header_format = self._workbook.add_format({
            'align':        "center", 
            'valign':       "vcenter", 
            'bold':         1, 
            'font_size':    12, 
            'border':       1, 
            'font_color':   'white', 
            'bg_color':     "#244062"})

        title_row = offset_y

        self._worksheet[sheet_index].set_row(title_row, 18)
        self._worksheet[sheet_index].set_column(offset_x + 0, offset_x + 0, 15)  # Player names
        self._worksheet[sheet_index].set_column(offset_x + 1, offset_x + 1, 22)  # Damage Done (DPS)
        self._worksheet[sheet_index].set_column(offset_x + 2, offset_x + 2, 18)  # Damage Taken
        self._worksheet[sheet_index].set_column(offset_x + 3, offset_x + 3, 8)   # Pot #1
        self._worksheet[sheet_index].set_column(offset_x + 4, offset_x + 4, 8)   # Pot #2
        self._worksheet[sheet_index].set_column(offset_x + 5, offset_x + 5, 9)   # Deaths

        self._worksheet[sheet_index].write(title_row, offset_x + 1, "Damage Taken (DPS)",   default_header_format)
        self._worksheet[sheet_index].write(title_row, offset_x + 2, "Damage Done",          default_header_format)
        self._worksheet[sheet_index].write(title_row, offset_x + 3, "Pot #1",               default_header_format)
        self._worksheet[sheet_index].write(title_row, offset_x + 4, "Pot #2",               default_header_format)
        self._worksheet[sheet_index].write(title_row, offset_x + 5, "Deaths",               default_header_format)

    def writeEncounterStats(self, difficulty, offset_x, offset_y):
        self._logger.log("\nCreating encounter statistic tables")
        self._logger.log("-----------------------------")

        cell_poty_format = self._workbook.add_format({'align': 'center', 'bg_color': "#92D050"})
        cell_potn_format = self._workbook.add_format({'align': 'center', 'bg_color': "#FF5757"})
        cell_stat_format_odd  = self._workbook.add_format({'align': 'center', 'bg_color': "#D9D9D9"})
        cell_stat_format_even = self._workbook.add_format({'align': 'center', 'bg_color': "#FFFFFF"})

        i = 0

        for encounter in self._zone_info['encounters']:
            self._logger.log("  " + self._shrinkBossName(encounter['name']))
            
            sheet_index = _WS_BOSS_INDEX_INIT + i
            i += 1

            self._writeBossStatsRow(sheet_index, offset_x, offset_y)
            
            player_row = offset_y + 1
            for name, player in sorted(self._player_data.items()):
                for fightID, fight_data in player['fights'].items():
                    # Skip unattended fights
                    if fight_data == None:
                        continue

                    if fight_data['kill'] == False or fight_data['boss'] != encounter['boss']:
                        continue

                    name_bg_color = "#323232" if (player_row % 2 == 0) else "#262626"

                    cell_name_format = self._workbook.add_format({'bold': 1, 'bg_color': "#323232", 'font_color': CLASS_COLOR[player['class']], 'bg_color': name_bg_color})
                    cell_stat_format = cell_stat_format_odd if (player_row % 2 == 0) else cell_stat_format_even
                    cell_poty_format = self._workbook.add_format({'align': 'center', 'bg_color': "#92D050"})
                    cell_potn_format = self._workbook.add_format({'align': 'center', 'bg_color': "#FF5757"})

                    stat_column = offset_x + 1

                    # Player Name
                    self._worksheet[sheet_index].write(player_row, offset_x,        name,   cell_name_format)  
                    # Damage Done (DPS)
                    self._worksheet[sheet_index].write(player_row, stat_column + 0, '-',    cell_stat_format)
                    # Damage Taken
                    self._worksheet[sheet_index].write(player_row, stat_column + 1, '-',    cell_stat_format)  

                    # Potion #1
                    if fight_data['enhancements']['pot_1'] == True:
                        self._worksheet[sheet_index].write(player_row, stat_column + 2, 'x', cell_poty_format)  # Used
                    else:
                        self._worksheet[sheet_index].write(player_row, stat_column + 2, '-', cell_potn_format)  # Missing

                    # Potion #2
                    if fight_data['enhancements']['pot_2'] == True:
                        self._worksheet[sheet_index].write(player_row, stat_column + 3, 'x', cell_poty_format)  # Used
                    else:
                        self._worksheet[sheet_index].write(player_row, stat_column + 3, '-', cell_potn_format)  # Missing

                    # Deaths
                    death_count = len(fight_data['deaths'])
                    if death_count == 0:
                        self._worksheet[sheet_index].write(player_row, stat_column + 4, '-',         cell_stat_format)  # None
                    else:
                        self._worksheet[sheet_index].write(player_row, stat_column + 4, death_count, cell_stat_format)  # Count



                    player_row += 1


