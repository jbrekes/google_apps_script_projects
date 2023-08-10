function updateTeamsAndInjuries(){
  var spreadsheet_id = NHLConfigs.spreadsheetID();

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
};
  
function getInjuredPlayers() {
  var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
  var url = 'https://statsapi.web.nhl.com';
  var teams_endpoint = '/api/v1/teams';

  function fetchJSON(endpoint) {
    var response = UrlFetchApp.fetch(url + endpoint);
    return JSON.parse(response.getContentText());
  }

  // Get Teams Full Names
  var teamsData = fetchJSON(teams_endpoint);
  var teams = teamsData.teams.map(function(teamData) {
    return teamData.name;
  });

  // Scrape Injuries Data
  var injuries_url = 'https://www.cbssports.com/nhl/injuries/';
  var response = UrlFetchApp.fetch(injuries_url).getContentText();
  var $ = Cheerio.load(response);

  // Extract Data
  var injured_players = [];

  $(".TableBase").each(function(i, element) {
    var t_name = $(element).find('.TeamName a').text();
    var t_name_fix = t_name;

    // Replace Scraped team name with the original version
    var fixedTeam = teams.find(function(team) {
      return team.includes(t_name);
    });

    if (fixedTeam) {
      t_name_fix = fixedTeam;
    }

    var players_names = [];
    $(element).find('.TableBase-bodyTd .CellPlayerName--long').each(function(i, el) {
      players_names.push($(el).text());
    });

    var players_positions = [];
    $(element).find('.TableBase-bodyTd').each(function(j, pos) {
      var pos_temp = $(pos).text().trim();
      if (pos_temp.length <= 2) {
        players_positions.push(pos_temp);
      }
    });

    var last_update = [];
    $(element).find('.TableBase-bodyTd .CellGameDate').each(function(k, up) {
      last_update.push($(up).text().trim());
    });

    var injury_status = [];
    $(element).find('.TableBase-bodyTr').each(function(l, stat) {
      var status = $(stat).find('.TableBase-bodyTd').last().text().trim();
      injury_status.push(status);
    });

    var injury_type = [];
    $(element).find('.TableBase-bodyTd').next().next().next().each(function(m, inj) {
      var inj_temp = $(inj).text().trim();
      if (injury_status.indexOf(inj_temp) === -1) {
        injury_type.push(inj_temp);
      }
    });

    for (var i = 0; i < players_names.length; i++) {
      var injured_temp = {
        'process_dt': process_dt,
        'team_name': t_name_fix,
        'player_name': players_names[i],
        'player_position': players_positions[i],
        'injury_name': injury_type[i],
        'injury_status': injury_status[i],
        'last_update': last_update[i]
      };

      injured_players.push(injured_temp);
    }
  });

  return injured_players;
};
  
function getTeamsData() {
  var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
  var url = 'https://statsapi.web.nhl.com';
  var teams_endpoint = '/api/v1/teams?expand=team.stats';
  var standings_endpoint = '/api/v1/standings';

  function fetchJSON(endpoint) {
    var response = UrlFetchApp.fetch(url + endpoint);
    return JSON.parse(response.getContentText());
  }

  // Get Main Team Data
  var teamsData = fetchJSON(teams_endpoint);
  var standingsData = fetchJSON(standings_endpoint);

  // Extract team data
  var teams = teamsData.teams.map(function(teamData) {
    var team = {
      'process_dt': process_dt,
      'team_id': teamData.id,
      'team_abbreviation': teamData.abbreviation,
      'team_name': teamData.name,
      'division_id': teamData.division.id,
      'division_name': teamData.division.name,
      'division_abbreviation': teamData.division.abbreviation,
      'conference_id': teamData.conference.id,
      'conference_name': teamData.conference.name
    };

    var stats = teamData.teamStats.find(function(stats) {
      return stats.type.displayName === 'statsSingleSeason';
    });

    if (stats) {
      Object.assign(team, stats.splits[0].stat);
    }

    return team;
  });

  // Extract standings data
  var standings = [];
  standingsData.records.forEach(function(record) {
    var div_id = record.division.id;
    var div_name = record.division.name;

    record.teamRecords.forEach(function(teamrec) {
      var team = {
        'team_id': teamrec.team.id,
        'team_name': teamrec.team.name,
        'division_id': div_id,
        'division_name': div_name,
        'league_record_wins': teamrec.leagueRecord.wins,
        'league_record_losses': teamrec.leagueRecord.losses,
        'league_record_ot': teamrec.leagueRecord.ot,
        'regulationWins': teamrec.regulationWins,
        'goalsAgainst': teamrec.goalsAgainst,
        'goalsScored': teamrec.goalsScored,
        'pts': teamrec.points,
        'league_rank': teamrec.leagueRank,
        'league_l10_rank': teamrec.leagueL10Rank,
        'divisionRank': teamrec.divisionRank,
        'division_l10_rank': teamrec.divisionL10Rank,
        'conferenceRank': teamrec.conferenceRank,
        'conference_l10_rank': teamrec.conferenceL10Rank,
        'streak_type': teamrec.streak.streakType,
        'streak_number': teamrec.streak.streakNumber,
        'streak_code': teamrec.streak.streakCode
      };
      standings.push(team);
    });
  });

  // Add Standings to Main Data
  teams.forEach(function(team) {
    var matchingStanding = standings.find(function(standing) {
      return standing.team_id === team.team_id;
    });

    if (matchingStanding) {
      Object.assign(team, matchingStanding);
    }
  });

  // Calculate Penalty Mins for each team and add it to teams data
  var roster_endpoint = '/api/v1/teams?expand=team.roster';
  var rosterData = fetchJSON(roster_endpoint);

  var pim_by_team = rosterData.teams.map(function(temp) {
    var team_name = temp.name;
    var total_mins = temp.roster.roster.reduce(function(total, player) {
      var playerStats = fetchJSON(player.person.link + '/stats?stats=statsSingleSeason');
      var penaltyMinutes = playerStats.stats[0]?.splits[0]?.stat?.penaltyMinutes || 0;
      return total + parseInt(penaltyMinutes);
    }, 0);

    return {
      'team_name': team_name,
      'pim': total_mins
    };
  });

  // Add Penalty Minutes to Team Stats
  teams.forEach(function(team) {
    var matchingPIM = pim_by_team.find(function(pim) {
      return pim.team_name === team.team_name;
    });

    if (matchingPIM) {
      team.pim = matchingPIM.pim;
    }
  });

  return teams;
};

function startingGoaliesScrap() {
  var process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
  var year = new Date().getFullYear();
  var sg_url = 'https://www.rotowire.com/hockey/starting-goalies.php?view=teams';
  var sg_response = UrlFetchApp.fetch(sg_url).getContentText();

  var $ = Cheerio.load(sg_response);

  function getFormattedDate(dateString) {
    if (dateString === 'Today') {
      return process_dt;
    } else {
      var temp = new Date(Date.parse(dateString));
      temp = Utilities.formatDate(temp, "GMT", "YYYY-MM-dd").replace(/^.{4}/, year);
      return temp;
    }
  }

  var dates = $('.proj-day-head').map(function(i, dt) {
    return getFormattedDate($(dt).text().trim().substring(3));
  }).get();

  // Fix for changing year
  for (var i = 0; i < (dates.length - 1); i++) {
    if (dates[i] > dates[i + 1]) {
      dates[i + 1] = dates[i + 1].replace(/^.{4}/, year + 1);
    }
  }

  var gk_data = $('.goalies-inner-row').map(function(i, item) {
    var gk_names = [];
    var gk_status = [];

    $(item).find('.goalie-item').each(function(j, temp) {
      gk_names.push($(temp).find('a').text().trim());
      gk_status.push($(temp).find('.sm-text').last().text().trim());
    });

    return gk_names.map(function(gk_name, i) {
      var gk = {
        'gk_date': dates[i],
        'gk_name': gk_name,
        'gk_status': gk_status[i]
      };
      return gk['gk_name'] !== "" ? gk : null;
    });
  }).get().flat();

  return gk_data;
};

function gamesSchedule() {
  const gk_scrap = startingGoaliesScrap();
  const process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
  const sch_url = 'https://statsapi.web.nhl.com';
  const sch_endpoint = '/api/v1/schedule';

  const days_back = 5;
  const days_forward = 60;

  const today = new Date();
  const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  let start_dt = Utilities.formatDate(new Date(today.getTime() - MILLIS_PER_DAY * days_back), "GMT", "YYYY-MM-dd");
  const end_dt = Utilities.formatDate(new Date(today.getTime() + MILLIS_PER_DAY * days_forward), "GMT", "YYYY-MM-dd");

  const games_sch = [];

  while (start_dt < end_dt) {
    let end_dt_temp = new Date(start_dt);
    end_dt_temp = Utilities.formatDate(new Date(end_dt_temp.getTime() + MILLIS_PER_DAY * 29), "GMT", "YYYY-MM-dd");
    if (end_dt_temp > end_dt) {
      end_dt_temp = end_dt;
    }

    const request_url = `${sch_url}${sch_endpoint}?startDate=${start_dt}&endDate=${end_dt_temp}`;
    const response = UrlFetchApp.fetch(request_url);
    const json = response.getContentText();
    const data = JSON.parse(json);

    for (const date of data.dates) {
      const d = date.date;

      for (const game of date.games) {
        const game_temp = {
          game_pk: String(game.gamePk),
          game_type: game.gameType,
          game_type_desc: game.gameType === 'R' ? 'Regular Season' :
                          game.gameType === 'P' ? 'Playoffs' :
                          game.gameType === 'PR' ? 'Pre-season' : 'Other',
          game_season: game.season,
          game_date: d,
          game_date_utc: game.gameDate,
          away_team_id: game.teams.away.team.id,
          away_team_name: game.teams.away.team.name,
          away_team_score: game.teams.away.score,
          home_team_id: game.teams.home.team.id,
          home_team_name: game.teams.home.team.name,
          home_team_score: game.teams.home.score
        };

        games_sch.push(game_temp);
      }
    }

    start_dt = new Date(start_dt);
    start_dt = Utilities.formatDate(new Date(start_dt.getTime() + MILLIS_PER_DAY * 30), "GMT", "YYYY-MM-dd");
  }

  for (const game of games_sch) {
    game.process_dt = process_dt;
  }

  return games_sch;
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

function updateScheduleToSheet() {
  const games_sch = goalkeepersByTeam();
  const process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
  const days_back = 5;
  const today = new Date();
  const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  let start_dt = Utilities.formatDate(new Date(today.getTime() - MILLIS_PER_DAY * days_back), "GMT", "YYYY-MM-dd");
  const spreadsheet_id = NHLConfigs.spreadsheetID();
  const sheet = SpreadsheetApp.openById(spreadsheet_id).getSheetByName("game_schedule");

  const c = 35; // Object.keys(games_sch[0]).length;
  const r = games_sch.length;

  // Delete most recent data
  const rows = sheet.getDataRange();
  const numRows = rows.getNumRows();
  const values = rows.getValues();

  let rowsDeleted = 0;
  for (let i = 1; i <= numRows - 1; i++) {
    const row = values[i];
    let row_date = row[1];
    row_date = Utilities.formatDate(row_date, 'GMT', 'YYYY-MM-dd');

    if (row_date >= start_dt) {
      sheet.deleteRow((parseInt(i) + 1) - rowsDeleted);
      rowsDeleted++;
    }
  }

  const last_row = sheet.getLastRow();
  const range = sheet.getRange(last_row + 1, 1, r, c);

  const table = games_sch.map(function(games_sch) {
    return [
      games_sch.process_dt, games_sch.game_date, games_sch.game_date_utc, games_sch.game_pk,
      games_sch.game_type, games_sch.game_type_desc, games_sch.game_season, games_sch.away_team_id,
      games_sch.away_team_name, games_sch.away_team_score, games_sch.home_team_id, games_sch.home_team_name,
      games_sch.home_team_score, games_sch.away_gk_name, games_sch.away_gk_id, games_sch.away_gk_games,
      games_sch.away_gk_wins, games_sch.away_gk_losses, games_sch.away_gk_ot, games_sch.away_gk_ties,
      games_sch.away_gk_shutouts, games_sch.away_gk_saves, games_sch.away_gk_gaa, games_sch.away_gk_save_percentage,
      games_sch.home_gk_name, games_sch.home_gk_id, games_sch.home_gk_games, games_sch.home_gk_wins,
      games_sch.home_gk_losses, games_sch.home_gk_ot, games_sch.home_gk_ties, games_sch.home_gk_shutouts,
      games_sch.home_gk_saves, games_sch.home_gk_gaa, games_sch.home_gk_save_percentage
    ];
  });

  range.setValues(table);
};

function getOdds_v1() {
  // API details
  const api_key = NHLConfigs.getApiKey();
  const market = 'h2h,spreads,totals';
  const regions = 'us';
  
  // Construct the API URL
  const url = `https://odds.p.rapidapi.com/v4/sports/icehockey_nhl/odds?regions=${regions}&oddsFormat=decimal&markets=${market}&dateFormat=iso`;

  // Set API headers
  const headers = {
    "X-RapidAPI-Key": api_key,
    "X-RapidAPI-Host": "odds.p.rapidapi.com"
  };

  // Configure API request options
  const options = {
    method: "GET",
    headers: headers
  };

  // Fetch odds data from the API
  const response = UrlFetchApp.fetch(url, options);
  const json = response.getContentText();
  const data = JSON.parse(json);

  // Prepare variables
  const process_dt = Utilities.formatDate(new Date(), "GMT", "YYYY-MM-dd");
  const next_games_odds = [];
  const bookmakers = ['SugarHouse', 'BetUS', 'Unibet', 'Betrivers', 'William Hill (US)', 'DraftKings', 'FanDuel', 'Bovada', 'FOX Bet'];

  // Extract relevant odds data
  for (const gameData of data) {
    const game_time = gameData['commence_time'];
    const home_team = gameData['home_team'];
    const away_team = gameData['away_team'];

    for (const bookmaker of gameData['bookmakers']) {
      const mkTitle = bookmaker['title'];

      if (bookmakers.includes(mkTitle) && bookmaker['markets'].length >= 3) {
        const marketData = bookmaker['markets'];

        const game = {
          'process_dt': process_dt,
          'game_time': game_time,
          'home_team': home_team,
          'away_team': away_team,
          'bookmaker_name': mkTitle,
          'h2h_home_price': marketData[0]['outcomes'][0]['price'],
          'h2h_away_price': marketData[0]['outcomes'][1]['price'],
          'spreads_home_price': marketData[1]['outcomes'][0]['price'],
          'spreads_home_points': marketData[1]['outcomes'][0]['point'],
          'spreads_away_price': marketData[1]['outcomes'][1]['price'],
          'spreads_away_points': marketData[1]['outcomes'][1]['point'],
          'totals_over_price': marketData[2]['outcomes'][0]['price'],
          'totals_over_points': marketData[2]['outcomes'][0]['point'],
          'totals_under_price': marketData[2]['outcomes'][1]['price'],
          'totals_under_points': marketData[2]['outcomes'][1]['point']
        };

        next_games_odds.push(game);
      }
    }
  }

  // Spreadsheet details
  const spreadsheet_id = NHLConfigs.spreadsheetID();
  const odds_sheet = SpreadsheetApp.openById(spreadsheet_id).getSheetByName("odds_1");
  const cols = Object.keys(next_games_odds[0]).length;
  const rorw = next_games_odds.length;

  // Clear previous information
  const del_range = odds_sheet.getRange(2, 1, 300, 20);
  del_range.clear();

  // Get a range for new data
  const range = odds_sheet.getRange(2, 1, rorw, cols);

  // Prepare data for insertion
  const table = next_games_odds.map(function(gameData) {
    return [
      gameData.process_dt, gameData.game_time, gameData.home_team, gameData.away_team,
      gameData.bookmaker_name, gameData.h2h_home_price, gameData.h2h_away_price,
      gameData.spreads_home_points, gameData.spreads_home_price, gameData.spreads_away_points,
      gameData.spreads_away_price, gameData.totals_over_points, gameData.totals_over_price,
      gameData.totals_under_points, gameData.totals_under_price
    ];
  });

  // Insert values into the spreadsheet
  range.setValues(table);
};
