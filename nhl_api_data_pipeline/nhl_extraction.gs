function updateTeamsAndInjuries(){
    var spreadsheet_id = '1a2NpAaysxopZlw93Nwbx9rQxD-GMbV3_vbeAZbZxCYc';
  
    // Update Teams Data
    var teams = getTeamsData();
    var teams_sheet = SpreadsheetApp.openById(spreadsheet_id).getSheetByName("teams");
  
    // Fix some format issues
    teams.forEach(function(dict){
      dict.ptPctg = dict.ptPctg / 100
      dict.powerPlayPercentage = dict.powerPlayPercentage / 100
      dict.penaltyKillPercentage = dict.penaltyKillPercentage / 100
      dict.faceOffWinPercentage = dict.faceOffWinPercentage / 100
      dict.shootingPctg = dict.shootingPctg / 100
      //dict.setType({})
    });
  
    // Get number of rows and columns I need
    var t_cols = Object.keys(teams[0]).length;
    var t_rorw = teams.length;
  
    // Delete Previous Information (just to be sure)
    var t_del_range = teams_sheet.getRange(2,1,70,70);
    t_del_range.clear();
  
    // Get a range for the second row
    var t_range = teams_sheet.getRange(2,1,t_rorw,t_cols);
  
    // Create a list of lists that represents the table:
    var t_table = teams.map(function(teams){
      return [teams.process_dt, teams.team_id, teams.team_name, teams.team_abbreviation, teams.division_id, teams.division_name, teams.division_abbreviation, teams.conference_id, teams.conference_name, teams.gamesPlayed, teams.wins, teams.losses, teams.ot, teams.pts, teams.ptPctg, teams.goalsPerGame, teams.goalsAgainstPerGame, teams.evGGARatio, teams.powerPlayPercentage, teams.powerPlayGoals, teams.powerPlayGoalsAgainst, teams.powerPlayOpportunities, teams.penaltyKillPercentage, teams.shotsPerGame, teams.shotsAllowed, teams.winScoreFirst, teams.winOppScoreFirst, teams.winLeadFirstPer, teams.winLeadSecondPer, teams.winOutshootOpp, teams.winOutshotByOpp, teams.faceOffsTaken, teams.faceOffsWon, teams.faceOffsLost, teams.faceOffWinPercentage, teams.shootingPctg, teams.savePctg, teams.league_record_wins, teams.league_record_losses, teams.league_record_ot, teams.regulationWins, teams.goalsAgainst, teams.goalsScored, teams.league_rank, teams.league_l10_rank, teams.divisionRank, teams.division_l10_rank, teams.conferenceRank, teams.conference_l10_rank, teams.streak_type, teams.streak_number, teams.streak_code, teams.pim] 
    })
  
    // Insert values into Sheet
    t_range.setValues(t_table)
  
    var team_finish_milestone = new Date();
    Logger.log('Finish Team Update');
    Logger.log(team_finish_milestone);
    Utilities.sleep(1000); // Pause for 1 second
  
    // Update Injuries
    var injured_players = getInjuredPlayers();
    var injuries_sheet = SpreadsheetApp.openById(spreadsheet_id).getSheetByName("current_injuries");
    var in_col = Object.keys(injured_players[0]).length
    var in_rows = injured_players.length 
  
    // Delete Previous Information (just to be sure)
    var i_del_range = injuries_sheet.getRange(2,1,500,10);
    i_del_range.clear();
  
    // Get a range for the second row
    var i_range = injuries_sheet.getRange(2,1,in_rows,in_col);
  
    // Create a list of lists that represents the table:
    var inj_table = injured_players.map(function(injured_players){
      return [injured_players.process_dt, injured_players.team_name, injured_players.player_name, injured_players.player_position, injured_players.injury_name, injured_players.injury_status, injured_players.last_update] 
    });
  
    // Insert values into Sheet
    i_range.setValues(inj_table);
  }
  
  function getInjuredPlayers(){
    var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
    // var spreadsheet_id = '1a2NpAaysxopZlw93Nwbx9rQxD-GMbV3_vbeAZbZxCYc';
    // var sheet = SpreadsheetApp.openById(spreadsheet_id).getSheetByName("current_injuries");
  
    //Get Teams Full Names
    var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
    var url = 'https://statsapi.web.nhl.com'
    var teams_endpoint = '/api/v1/teams'
  
    // Get Main Team Data
    var response = UrlFetchApp.fetch(url + teams_endpoint);
  
    var json = response.getContentText();
    var data = JSON.parse(json);
  
    var teams = [];
    for (var i in data['teams']) {
      // Get team base info
      var team_name = data['teams'][i]['name'];
      teams.push(team_name);
    }
  
    // Scrape Injuries Data
  
    var injuries_url = 'https://www.cbssports.com/nhl/injuries/';
    var response = UrlFetchApp.fetch(injuries_url).getContentText();
  
    var $ = Cheerio.load(response)
  
    // Extract Data
  
    var injured_players = [];
  
    var teams_names_scrap = $(".TableBase").each((i,element) => {
      var t_name = $(element).find('.TeamName').find('a').text();
  
      // Replace Scraped team name with the original version
      if (t_name === 'Montreal'){
        var t_name_fix = ['Montréal Canadiens'] 
      } else if (t_name === 'N.Y. Islanders'){
          var t_name_fix = ['New York Islanders']
      } else if (t_name === 'N.Y. Rangers'){
          var t_name_fix = ['New York Rangers']
      } else {
          for (var i = 0; i < teams.length; i++){
            if (teams[i].includes(t_name)){
              var t_name_fix = [teams[i]];
            }
          }
      }
  
      var players_names = []
      var player_names = $(element).find('.TableBase-bodyTd .CellPlayerName--long').each((i, el) => {
        players_names.push($(el).text())
      });
  
      var players_positions = []
      var player_pos = $(element).find('.TableBase-bodyTd').each((j,pos) => {
        var pos_temp = $(pos).text().trim();
  
        if (pos_temp.length <= 2){
          players_positions.push(pos_temp)
        } else {
          // Nothing
        }
      });
  
      var last_update = []
      var updates = $(element).find('.TableBase-bodyTd .CellGameDate').each((k, up) => {
        last_update.push($(up).text().trim())
      });
      
      var injury_status = []
      var inj_stat = $(element).find('.TableBase-bodyTr').each((l,stat) => {
        status = $(stat).find('.TableBase-bodyTd').last().text().trim()
        injury_status.push(status)
      });
  
      var injury_type = []
      var inj_type = $(element).find('.TableBase-bodyTd').next().next().next().each((m,inj) => {
        var inj_temp = $(inj).text().trim();
  
        if (injury_status.indexOf(inj_temp) === -1){
          injury_type.push(inj_temp)
        } else {
          // Nothing
        }
      });
  
      for (var i = 0; i < players_names.length; i ++){
        var injured_temp = {
          'process_dt': process_dt,
          'team_name': t_name_fix[0],
          'player_name': players_names[i],
          'player_position': players_positions[i],
          'injury_name': injury_type[i],
          'injury_status': injury_status[i],
          'last_update': last_update[i]
        }
  
        injured_players.push(injured_temp)
  
      }
    });
    return injured_players;
    // // Add all data to the spreadsheet
    // var c = Object.keys(injured_players[0]).length
    // var r = injured_players.length
  
    // // Delete Previous Information (just to be sure)
    // var del_range = sheet.getRange(2,1,500,10);
    // del_range.clear();
  
    // // Get a range for the second row
    // var range = sheet.getRange(2,1,r,c);
  
    // // Create a list of lists that represents the table:
    // var table = injured_players.map(function(injured_players){
    //   return [injured_players.process_dt, injured_players.team_name, injured_players.player_name, injured_players.player_position, injured_players.injury_name, injured_players.injury_status, injured_players.last_update] 
    // })
  
    // // Insert values into Sheet
    // range.setValues(table)
  };
  
  function getTeamsData() {
    var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
    var url = 'https://statsapi.web.nhl.com'
    var teams_endpoint = '/api/v1/teams?expand=team.stats'
    var standings_endpoint = '/api/v1/standings'
  
    // Get Main Team Data
    var response = UrlFetchApp.fetch(url + teams_endpoint);
  
    var json = response.getContentText();
    var data = JSON.parse(json);
  
    var teams = [];
    for (var i in data['teams']) {
      // Get team base info
      var team = {
        'process_dt': process_dt,
        'team_id': data['teams'][i]['id'],
        'team_abbreviation': data['teams'][i]['abbreviation'],
        'team_name': data['teams'][i]['name'],
        'division_id': data['teams'][i]['division']['id'],
        'division_name': data['teams'][i]['division']['name'],
        'division_abbreviation': data['teams'][i]['division']['abbreviation'],
        'conference_id': data['teams'][i]['conference']['id'],
        'conference_name': data['teams'][i]['conference']['name']     
      }
  
      // Get Stats
      var stats_temp = [];
      for (var j in data['teams'][i]['teamStats']) {
        if (data['teams'][i]['teamStats'][j]['type']['displayName'] == 'statsSingleSeason') {
          stats_temp.push(data['teams'][i]['teamStats'][j]['splits'][0]['stat']);
        }
      }
  
      for (var key in stats_temp[0]) {
        team[key] = stats_temp[0][key];
      }
  
      teams.push(team);
    }
  
    // Get Standings
    var st_response = UrlFetchApp.fetch(url + standings_endpoint);
  
    var st_json = st_response.getContentText();
    var st_data = JSON.parse(st_json);
  
    var standings = [];
  
    for (var i in st_data['records']) {
      var record = st_data['records'][i];
      var div_id = record['division']['id'];
      var div_name = record['division']['name'];
  
      for (var j in record['teamRecords']) {
        var teamrec = record['teamRecords'][j]
        var team = {
          'team_id': teamrec['team']['id'],
          'team_name': teamrec['team']['name'],
          'division_id': div_id,
          'division_name': div_name,
          'league_record_wins': teamrec['leagueRecord']['wins'],
          'league_record_losses': teamrec['leagueRecord']['losses'],
          'league_record_ot': teamrec['leagueRecord']['ot'],
          'regulationWins': teamrec['regulationWins'],
          'goalsAgainst': teamrec['goalsAgainst'],
          'goalsScored': teamrec['goalsScored'],
          'pts': teamrec['points'],
          'league_rank': teamrec['leagueRank'],
          'league_l10_rank': teamrec['leagueL10Rank'],
          'divisionRank': teamrec['divisionRank'],
          'division_l10_rank': teamrec['divisionL10Rank'],
          'conferenceRank': teamrec['conferenceRank'],
          'conference_l10_rank': teamrec['conferenceL10Rank'],
          'streak_type': teamrec['streak']['streakType'],
          'streak_number': teamrec['streak']['streakNumber'],
          'streak_code': teamrec['streak']['streakCode']
        }
        standings.push(team);
      }
    }
  
    // Add Standings to Main Data
    for (var i in teams){
      var team_id = teams[i]['team_id']
  
      // Filter the Standings for the selected team
      var standings_temp = standings.filter(function(standings){
        return standings.team_id == team_id
      });
  
      // Combine Standings dict and teams dict
      teams[i] = Object.assign({}, teams[i], standings_temp[0])
    }
  
    // Calculate Penalty Mins for each team and add it to teams data
    var roster_endpoint = '/api/v1/teams?expand=team.roster';
    var roster_response = UrlFetchApp.fetch(url + roster_endpoint)
  
    var roster_json = roster_response.getContentText();
    var roster_data = JSON.parse(roster_json);
  
    // First calculate PIM for each Team
    var pim_by_team = [];
  
    for (var i in roster_data['teams']){
      var temp = roster_data['teams'][i];
      var team_name = temp['name'];
      var roster = temp['roster']['roster'];
      var total_mins = 0;
  
      for (var p in roster){
        var player_request = JSON.parse(UrlFetchApp.fetch(url + roster[p]['person']['link'] + '/stats?stats=statsSingleSeason').getContentText())
  
        if ('stats' in player_request && player_request['stats'].length > 0 && 'splits' in player_request['stats'][0] && player_request['stats'][0]['splits'].length > 0 && 'penaltyMinutes' in player_request['stats'][0]['splits'][0]['stat']) {
          total_mins += parseInt(player_request['stats'][0]['splits'][0]['stat']['penaltyMinutes']);
        } else {
          total_mins += 0;
        }
      }
      var pim = {
        'team_name': team_name,
        'pim': total_mins
      };
  
      pim_by_team.push(pim)
    };
  
    // Now Add it to the Team stats
    for (var i in teams){
      var team_name = teams[i]['team_name'];
      var pim_filtered = pim_by_team.filter(function(pim_by_team){
        return pim_by_team.team_name == team_name
      });
  
      teams[i]['pim'] = pim_filtered[0]['pim'];
    };
  
    return teams; 
  
    // // Fix some format issues
    // teams.forEach(function(dict){
    //   dict.ptPctg = dict.ptPctg / 100
    //   dict.powerPlayPercentage = dict.powerPlayPercentage / 100
    //   dict.penaltyKillPercentage = dict.penaltyKillPercentage / 100
    //   dict.faceOffWinPercentage = dict.faceOffWinPercentage / 100
    //   dict.shootingPctg = dict.shootingPctg / 100
    //   //dict.setType({})
    // })
  
    // // Paste values into Google Sheets
    // var spreadsheet_id = '1a2NpAaysxopZlw93Nwbx9rQxD-GMbV3_vbeAZbZxCYc'
    // var sheet = SpreadsheetApp.openById(spreadsheet_id).getSheetByName("teams");
  
    // // Get number of rows and columns I need
    // var c = Object.keys(teams[0]).length
    // var r = teams.length
  
    // Logger.log(c);
    // Logger.log(r);
  
    // // Delete Previous Information (just to be sure)
    // var del_range = sheet.getRange(2,1,70,70);
    // del_range.clear();
  
    // // Get a range for the second row
    // var range = sheet.getRange(2,1,r,c);
  
    // // Create a list of lists that represents the table:
    // var table = teams.map(function(teams){
    //   return [teams.process_dt, teams.team_id, teams.team_name, teams.team_abbreviation, teams.division_id, teams.division_name, teams.division_abbreviation, teams.conference_id, teams.conference_name, teams.gamesPlayed, teams.wins, teams.losses, teams.ot, teams.pts, teams.ptPctg, teams.goalsPerGame, teams.goalsAgainstPerGame, teams.evGGARatio, teams.powerPlayPercentage, teams.powerPlayGoals, teams.powerPlayGoalsAgainst, teams.powerPlayOpportunities, teams.penaltyKillPercentage, teams.shotsPerGame, teams.shotsAllowed, teams.winScoreFirst, teams.winOppScoreFirst, teams.winLeadFirstPer, teams.winLeadSecondPer, teams.winOutshootOpp, teams.winOutshotByOpp, teams.faceOffsTaken, teams.faceOffsWon, teams.faceOffsLost, teams.faceOffWinPercentage, teams.shootingPctg, teams.savePctg, teams.league_record_wins, teams.league_record_losses, teams.league_record_ot, teams.regulationWins, teams.goalsAgainst, teams.goalsScored, teams.league_rank, teams.league_l10_rank, teams.divisionRank, teams.division_l10_rank, teams.conferenceRank, teams.conference_l10_rank, teams.streak_type, teams.streak_number, teams.streak_code, teams.pim] 
    // })
  
    // // Insert values into Sheet
    // range.setValues(table)
  };
  
  function startingGoaliesScrap(){
    var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
    var year = new Date().getFullYear();
    var sg_url = 'https://www.rotowire.com/hockey/starting-goalies.php?view=teams'
    var sg_response = UrlFetchApp.fetch(sg_url).getContentText();
  
    var $ = Cheerio.load(sg_response)
  
    // Get dates
    var dates = []
    $('.proj-day-head').each((i,dt) => {
      temp = $(dt).text();
  
      if (temp === 'Today'){
        dates.push(process_dt)
      } else{
        // Get just the short date
        temp = temp.substring(3)
        // Adapt format
        temp = new Date(Date.parse(temp));
        temp = Utilities.formatDate(temp, "GMT", "YYYY-MM-dd").replace(/^.{4}/, year);
        // Add to list
        dates.push(temp);
      };    
    });
  
    // Fix for when we change year
    for (var i = 0; i < (dates.length - 1); i++){
      if (dates[i] > dates[i+1]){
        dates[i+1] = dates[i+1].replace(/^.{4}/, year+1);
      } else {
        // Nothing
      };
    };
  
    // Schedule By team for next week
  
    // Goalie Name
    gk_data = []
  
    $('.goalies-inner-row').each((i,item) => {
  
      gk_names = []
      gk_status = []
  
      name = $(item).find('.goalie-item').each((j,temp) => {
        // Goalie Name
        gk_names.push($(temp).find('a').text().trim())
        // Goalie Status
        gk_status.push($(temp).find('.sm-text').last().text().trim())
      });
  
      for (var i = 0; i < gk_names.length; i ++){
        gk = {
          'gk_date': dates[i],
          'gk_name': gk_names[i],
          'gk_status': gk_status[i]
        }
        if (!(gk['gk_name'] === "")){
          gk_data.push(gk)
        } else {
          // Nothing
        }
      }
      
    });
    return gk_data
  };
  
  function gamesSchedule(){
    var gk_scrap = startingGoaliesScrap();
    var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
    var sch_url = 'https://statsapi.web.nhl.com';
    var sch_endpoint = '/api/v1/schedule';
  
    var days_back = 5
    var days_forward = 60
  
    var today = new Date();
    const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var start_dt = Utilities.formatDate(new Date(today.getTime() - MILLIS_PER_DAY * days_back), "GMT", "YYYY-MM-dd");
    var end_dt = Utilities.formatDate(new Date(today.getTime() + MILLIS_PER_DAY * days_forward), "GMT", "YYYY-MM-dd");
  
    games_sch = []
  
    while (start_dt < end_dt){
  
      // Create a temp end date so we process smaller ranges
      end_dt_temp = new Date(start_dt);
      end_dt_temp = Utilities.formatDate(new Date(end_dt_temp.getTime() + MILLIS_PER_DAY * 29), "GMT", "YYYY-MM-dd");
      if (end_dt_temp > end_dt){
        end_dt_temp = end_dt
      };
  
      var request_url = sch_url +  sch_endpoint + '?startDate=' + start_dt + '&endDate=' + end_dt_temp;
      var response = UrlFetchApp.fetch(request_url);
      var json = response.getContentText();
      var data = JSON.parse(json);
  
      for (var i in data['dates']){
        dt = data['dates'][i]
        d = dt['date']
  
        for (var game in dt['games']){
          g = dt['games'][game];
  
          var game_temp = {
            'game_pk': String(g['gamePk']),
            'game_type': g['gameType'],
            'game_type_desc': g['gameType'] == 'R' ? 'Regular Season' :
                              g['gameType'] == 'P' ? 'Playoffs' :
                              g['gameType'] == 'PR' ? 'Pre-season' : 'Other',
            'game_season': g['season'],
            'game_date': d,
            'game_date_utc': g['gameDate'],
            'away_team_id': g['teams']['away']['team']['id'],
            'away_team_name': g['teams']['away']['team']['name'],
            'away_team_score': g['teams']['away']['score'],
            'home_team_id': g['teams']['home']['team']['id'],
            'home_team_name': g['teams']['home']['team']['name'],
            'home_team_score': g['teams']['home']['score']
          }
  
          games_sch.push(game_temp);
        };
      };
  
      // Next Start Date
      start_dt = new Date(start_dt);
      start_dt = Utilities.formatDate(new Date(start_dt.getTime() + MILLIS_PER_DAY * 30), "GMT", "YYYY-MM-dd");
    };
    for (var i in games_sch){
      games_sch[i]['process_dt'] = process_dt;
    };
    return games_sch
  };
  
  function goalkeepersByTeam(){
    var starting_goalies = startingGoaliesScrap()
    var games_sch = gamesSchedule()
    var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
    var url = 'https://statsapi.web.nhl.com';
    var roster_endpoint = '/api/v1/teams?expand=team.roster';
  
    var response = UrlFetchApp.fetch(url + roster_endpoint);
    var json = response.getContentText();
    var data = JSON.parse(json);
  
    goalies = [];
  
    // Search for goalies basic data and id
    for (var i in data['teams']){
      team_data = data['teams'][i]
      roster = team_data['roster']['roster']
  
      for (var p in roster){
        var rs = roster[p];
  
        if (rs['position']['name'] === 'Goalie'){
          goalie = {
            'process_dt': process_dt,
            'team_id': String(team_data['id']),
            'team_name': team_data['name'],
            'team_abbreviation': team_data['abbreviation'],
            'goalie_name': rs['person']['fullName'],
            'goalie_id': String(rs['person']['id']),
            'goalie_link': rs['person']['link'],
            'jersey_number': String(rs['jerseyNumber'])
          };
  
          goalies.push(goalie)
        } else {
          // Nothing
        };
  
      };
    };
  
    // Add goalie stats
  
    for (var i in goalies){
      var goalie_stats_url = url + goalies[i]['goalie_link'] + '/stats?stats=statsSingleSeason';
  
      var goalie_stats_response = JSON.parse(UrlFetchApp.fetch(goalie_stats_url).getContentText());
      
      if ('stats' in goalie_stats_response &&
          goalie_stats_response['stats'].length > 0 &&
          'splits' in goalie_stats_response['stats'][0] &&
          goalie_stats_response['stats'][0]['splits'].length > 0) {
        var temp = goalie_stats_response['stats'][0]['splits'][0]['stat'];
  
        var goalie_stats = {
          'games': temp['games']  || null,
          'wins': temp['wins'] || null,
          'losses': temp['losses'] || null,
          'ot': temp['ot'] || null,
          'ties': temp['ties'] || null,
          'shutouts': temp['shutouts'] || null,
          'saves': temp['saves'] || null,
          'gaa': temp['goalAgainstAverage'] || null,
          'save_percentage': temp['savePercentage'] || null
          // etc.
        };
      } else {
        // handle the case where stats or splits is not present or is empty
        var goalie_stats = {
          'games': null,
          'wins': null,
          'losses': null,
          'ot': null,
          'ties': null,
          'shutouts': null,
          'saves': null,
          'gaa': null,
          'save_percentage': null
          // etc.
        };
      };
      goalies[i] = Object.assign(goalies[i], goalie_stats);
    };
  
    // Add Starting Goalkeepers and their stats to the scheduled games
    for (var i in games_sch){
      var game = games_sch[i];
      if (game['game_date'] < process_dt){
        var gk_endpoint = '/api/v1/game/' + game['game_pk'] + '/boxscore';
        var gk_request = JSON.parse(UrlFetchApp.fetch(url + gk_endpoint).getContentText());
        var locs = ['away', 'home'];
  
        for (var i = 0; i < locs.length; i++){
          var loc = String(locs[i]);
          var players = gk_request['teams'][loc]['players'];
          var main_goalie = {
              'gk_id': null,
              'gk_name': null,
              'gk_toi_final': 0,
              'gk_shots': 0
          };
          for (var player in players){
            var data = players[player];
            
            if (data['position']['name'] === 'Goalie'){
              var gk_id = data['person']['id'];
              var gk_name = data['person']['fullName'];
  
              try {
                // Search for the goalie with more time on ice or shots
                var gk_toi = data['stats']['goalieStats']['timeOnIce'];
                var gk_toi_dot_index = gk_toi.indexOf(':');
                var gk_toi_final = Number(gk_toi.substring(0,gk_toi_dot_index));
                var gk_shots = data['stats']['goalieStats']['shots'];
  
                if (gk_toi_final > main_goalie['gk_toi_final'] || gk_shots > main_goalie['gk_shots']){
                  main_goalie = {
                      'gk_id': gk_id,
                      'gk_name': gk_name,
                      'gk_toi_final': gk_toi_final,
                      'gk_shots': gk_shots
                  }
                }else{
                  // Nothing
                }
              } catch (e){
                // continue
              };
            }else{
              // Nothing
            };
          };
          try{
            var gk_req = JSON.parse(UrlFetchApp.fetch(url + '/api/v1/people/' + String(main_goalie['gk_id']) + '/stats?stats=statsSingleSeason&season=' + game['game_season']).getContentText());
            var temp = gk_req['stats'][0]['splits'][0]['stat'];
  
            var gk = {
              [`${loc}_gk_name`] : main_goalie['gk_name'],
              [`${loc}_gk_id`]: main_goalie['gk_id'],
              [`${loc}_gk_games`]: temp['games'] || null,
              [`${loc}_gk_wins`]: temp['wins'] || null,
              [`${loc}_gk_losses`]: temp['losses'] || null,
              [`${loc}_gk_ot`]: temp['ot'] || null,
              [`${loc}_gk_ties`]: temp['ties'] || null,
              [`${loc}_gk_shutouts`]: temp['shutouts'] || null,
              [`${loc}_gk_saves`]: temp['saves'] || null,
              [`${loc}_gk_gaa`]: temp['goalAgainstAverage'] || null,
              [`${loc}_gk_save_percentage`]: temp['savePercentage'] || null
            };
  
            game = Object.assign(game, gk);
  
          } catch (e){
            var gk = {
              [`${loc}_gk_name`]: main_goalie['gk_name'],
              [`${loc}_gk_id`]: main_goalie['gk_id'],
              [`${loc}_gk_games`]: null,
              [`${loc}_gk_wins`]: null,
              [`${loc}_gk_losses`]: null,
              [`${loc}_gk_ot`]: null,
              [`${loc}_gk_ties`]: null,
              [`${loc}_gk_shutouts`]: null,
              [`${loc}_gk_saves`]: null,
              [`${loc}_gk_gaa`]: null,
              [`${loc}_gk_save_percentage`]: null
            };
            game = Object.assign(game, gk);
          };
        };
      }else{
        var gk = {
          [`${loc}_gk_name`]: null,
          [`${loc}_gk_id`]: null,
          [`${loc}_gk_games`]: null,
          [`${loc}_gk_wins`]: null,
          [`${loc}_gk_losses`]: null,
          [`${loc}_gk_ot`]: null,
          [`${loc}_gk_ties`]: null,
          [`${loc}_gk_shutouts`]: null,
          [`${loc}_gk_saves`]: null,
          [`${loc}_gk_gaa`]: null,
          [`${loc}_gk_save_percentage`]: null
        };
        game = Object.assign(game, gk);
      };
    };
    
    // Now add goalkeepers for current and future games
    for (var i in games_sch){
      if (games_sch[i]['game_date'] >= process_dt){
        // Search for goalkeepers playing on that date
        var filtered_goalies_date = starting_goalies.filter(function(gl){
          return gl.gk_date == games_sch[i]['game_date'];
        });
        if (filtered_goalies_date.length > 0){
          for (var g in filtered_goalies_date){
            var temp_gk = filtered_goalies_date[g];
  
            var gk_stats = []
            for (var st in goalies){
              var temp = goalies[st];
  
              if (temp['goalie_name'] == temp_gk['gk_name']){
                gk_stats.push(temp)
              };
            };
  
            var gk_stats = gk_stats[0];
  
            // Set as home or away goalie
            try{
              if (games_sch[i]['home_team_name'] == gk_stats['team_name']){
                gk_home = {
                  'home_gk_name': gk_stats['goalie_name'],
                  'home_gk_id': gk_stats['goalie_id'],
                  'home_gk_games': gk_stats['games'],
                  'home_gk_wins': gk_stats['wins'],
                  'home_gk_losses': gk_stats['losses'],
                  'home_gk_ot': gk_stats['ot'],
                  'home_gk_ties': gk_stats['ties'],
                  'home_gk_shutouts': gk_stats['shutouts'],
                  'home_gk_saves': gk_stats['saves'],
                  'home_gk_gaa': gk_stats['gaa'],
                  'home_gk_save_percentage': gk_stats['save_percentage'],
                };
                games_sch[i] = Object.assign(games_sch[i], gk_home);
              } else if (games_sch[i]['away_team_name'] == gk_stats['team_name']){
                gk_away = {
                  'away_gk_name': gk_stats['goalie_name'],
                  'away_gk_id': gk_stats['goalie_id'],
                  'away_gk_games': gk_stats['games'],
                  'away_gk_wins': gk_stats['wins'],
                  'away_gk_losses': gk_stats['losses'],
                  'away_gk_ot': gk_stats['ot'],
                  'away_gk_ties': gk_stats['ties'],
                  'away_gk_shutouts': gk_stats['shutouts'],
                  'away_gk_saves': gk_stats['saves'],
                  'away_gk_gaa': gk_stats['gaa'],
                  'away_gk_save_percentage': gk_stats['save_percentage'],
                };
                games_sch[i] = Object.assign(games_sch[i], gk_away);
              } else{
                // Nothing
              };
            }catch (e) {
              continue
            };
          };
        }else{
          gk_home = {
            'home_gk_name': null,
            'home_gk_id': null,
            'home_gk_games': null,
            'home_gk_wins': null,
            'home_gk_losses': null,
            'home_gk_ot': null,
            'home_gk_ties': null,
            'home_gk_shutouts': null,
            'home_gk_saves': null,
            'home_gk_gaa': null,
            'home_gk_save_percentage': null,
          };
          gk_away = {
            'away_gk_name': null,
            'away_gk_id': null,
            'away_gk_games': null,
            'away_gk_wins': null,
            'away_gk_losses': null,
            'away_gk_ot': null,
            'away_gk_ties': null,
            'away_gk_shutouts': null,
            'away_gk_saves': null,
            'away_gk_gaa': null,
            'away_gk_save_percentage': null,
          };
          games_sch[i] = Object.assign(games_sch[i], gk_home);
          games_sch[i] = Object.assign(games_sch[i], gk_away);
        };
      }else{
        // Nothing (for game_date < process_dt)
      };
    };
    return games_sch
  };
  
  function updateScheduleToSheet(){
    var games_sch = goalkeepersByTeam();
    var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
    var days_back = 5;
    var today = new Date();
    const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var start_dt = Utilities.formatDate(new Date(today.getTime() - MILLIS_PER_DAY * days_back), "GMT", "YYYY-MM-dd");
    var spreadsheet_id = '1a2NpAaysxopZlw93Nwbx9rQxD-GMbV3_vbeAZbZxCYc';
    var sheet = SpreadsheetApp.openById(spreadsheet_id).getSheetByName("game_schedule");
  
    // Add all data to the spreadsheet. We will preserve the older data
    var c = 35 // Object.keys(games_sch[0]).length;
    var r = games_sch.length;
    Logger.log(c)
    Logger.log(r)
    Logger.log(start_dt)
  
    // Delete most recent data
    var rows = sheet.getDataRange();
    var numRows = rows.getNumRows();
    var values = rows.getValues();
  
    var rowsDeleted = 0;
    for (var i = 1; i <= numRows - 1; i++) {
      var row = values[i];
      var row_date = row[1];
      row_date = Utilities.formatDate(row_date, 'GMT', 'YYYY-MM-dd');
  
      if (row_date >= start_dt){
        sheet.deleteRow((parseInt(i)+1) - rowsDeleted);
        rowsDeleted++;
      };
    };
  
    var last_row = sheet.getLastRow();
  
    // Get a range for the second row
    var range = sheet.getRange(last_row+1,1,r,c);
  
    // Create a list of lists that represents the table:
    var table = games_sch.map(function(games_sch){
      return [games_sch.process_dt, games_sch.game_date, games_sch.game_date_utc, games_sch.game_pk, games_sch.game_type, games_sch.game_type_desc, games_sch.game_season,games_sch.away_team_id, games_sch.away_team_name, games_sch.away_team_score, games_sch.home_team_id, games_sch.home_team_name, games_sch.home_team_score, games_sch.away_gk_name, games_sch.away_gk_id, games_sch.away_gk_games, games_sch.away_gk_wins, games_sch.away_gk_losses, games_sch.away_gk_ot, games_sch.away_gk_ties, games_sch.away_gk_shutouts, games_sch.away_gk_saves, games_sch.away_gk_gaa, games_sch.away_gk_save_percentage, games_sch.home_gk_name, games_sch.home_gk_id, games_sch.home_gk_games, games_sch.home_gk_wins, games_sch.home_gk_losses, games_sch.home_gk_ot, games_sch.home_gk_ties, games_sch.home_gk_shutouts, games_sch.home_gk_saves, games_sch.home_gk_gaa, games_sch.home_gk_save_percentage] 
    })
  
    // Insert values into Sheet and remove dupplicates just in case
    range.setValues(table);
    Logger.log(sheet.getLastRow())
  };
  
  function getOdds_v1(){
    // Connect to API
    var api_key = 'accd61f86fmshb0388fa5577e465p1d7388jsn2d2e5cf0a3aa';
    var market = 'h2h,spreads,totals';
    var regions = 'us'
  
    var url = "https://odds.p.rapidapi.com/v4/sports/icehockey_nhl/odds" + '?regions=' + regions + '&oddsFormat=decimal' + '&markets=' + market + '&dateFormat=iso'
    var querystring = {
      "regions":"us",
      "oddsFormat":"decimal",
      "markets":"h2h,spreads,totals",
      "dateFormat":"iso"
    };
    var headers = {
      "X-RapidAPI-Key": "accd61f86fmshb0388fa5577e465p1d7388jsn2d2e5cf0a3aa",
      "X-RapidAPI-Host": "odds.p.rapidapi.com"
    };
    var options = {
      method: "GET",
      headers: headers
    };
  
    var response = UrlFetchApp.fetch(url,options);
    var json = response.getContentText();
    var data = JSON.parse(json);
  
    //  Extract odd data for some bookmakers
    var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
    var next_games_odds = [];
    var bookmakers = ['SugarHouse','BetUS','Unibet','Betrivers','William Hill (US)','DraftKings','FanDuel','Bovada','FOX Bet']
  
    for (var i in data){
      var game_time = data[i]['commence_time'];
      var home_team = data[i]['home_team'];
      var away_team = data[i]['away_team'];
  
      for (var j in data[i]['bookmakers']){
        var mk = data[i]['bookmakers'][j];
  
        if (bookmakers.indexOf(mk['title']) != -1){
          if (mk['markets'].length >= 3){
            var game = {
              'process_dt': process_dt,
              'game_time': game_time,
              'home_team': home_team,
              'away_team': away_team,
              'bookmaker_name': mk['title'],
              'h2h_home_price': mk['markets'][0]['outcomes'][0]['price'],
              'h2h_away_price': mk['markets'][0]['outcomes'][1]['price'],
              'spreads_home_price': mk['markets'][1]['outcomes'][0]['price'],
              'spreads_home_points': mk['markets'][1]['outcomes'][0]['point'],
              'spreads_away_price': mk['markets'][1]['outcomes'][1]['price'],
              'spreads_away_points': mk['markets'][1]['outcomes'][1]['point'],
              'totals_over_price': mk['markets'][2]['outcomes'][0]['price'],
              'totals_over_points': mk['markets'][2]['outcomes'][0]['point'],
              'totals_under_price': mk['markets'][2]['outcomes'][1]['price'],
              'totals_under_points': mk['markets'][2]['outcomes'][1]['point'],
            }
            next_games_odds.push(game)
          };
        };
  
      };
  
    };
  
    //  Add it to Google Sheets
    var spreadsheet_id = '1a2NpAaysxopZlw93Nwbx9rQxD-GMbV3_vbeAZbZxCYc';
    var odds_sheet = SpreadsheetApp.openById(spreadsheet_id).getSheetByName("odds_1");
  
    // Get number of rows and columns I need
    var cols = Object.keys(next_games_odds[0]).length;
    var rorw = next_games_odds.length;
  
    // Delete Previous Information (just to be sure)
    var del_range = odds_sheet.getRange(2,1,300,20);
    del_range.clear();
  
    // Get a range for the second row
    var range = odds_sheet.getRange(2,1,rorw,cols);
  
    // Create a list of lists that represents the table:
    var table = next_games_odds.map(function(next_games_odds){
      return [next_games_odds.process_dt, next_games_odds.game_time, next_games_odds.home_team, next_games_odds.away_team, next_games_odds.bookmaker_name, next_games_odds.h2h_home_price, next_games_odds.h2h_away_price, next_games_odds.spreads_home_points, next_games_odds.spreads_home_price, next_games_odds.spreads_away_points, next_games_odds.spreads_away_price, next_games_odds.totals_over_points, next_games_odds.totals_over_price, next_games_odds.totals_under_points, next_games_odds.totals_under_price] 
    })
  
    // Insert values into Sheet
    range.setValues(table)
  };