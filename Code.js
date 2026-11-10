function doGet(e) {
  const requested = e && e.parameter && e.parameter.page;
  const page = requested === 'games' ? 'Games' : (requested === 'admin' ? 'Admin' : 'Index');
  const titles = { Games: 'Game Results', Admin: 'Admin - Add Game', Index: 'League Table' };
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(titles[page])
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getLeagueTableData(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Full Raw data');

  if (!sheet) {
    throw new Error("Sheet 'Full Raw data' not found.");
  }

  const values = sheet.getDataRange().getValues();

  if (!values || values.length < 2) {
    return {
      rows: [],
      filters: { years: [], seasons: [], weeks: [] },
      selected: { year: '', season: '', week: '' }
    };
  }

  const headers = values[0].map(function (header) {
    return String(header || '').trim().toLowerCase();
  });

  const indexes = {
    year: headers.indexOf('year'),
    season: headers.indexOf('season'),
    week: headers.indexOf('week'),
    teams: headers.indexOf('teams'),
    players: headers.indexOf('players'),
    goals: headers.indexOf('goals'),
    bibsScore: findScoreColumn(headers, 'bibsscore'),
    nonBibsScore: findScoreColumn(headers, 'nonbibsscore')
  };

  const required = ['year', 'season', 'week', 'teams', 'players', 'goals'];
  const missing = required.filter(function (key) {
    return indexes[key] === -1;
  });

  if (missing.length) {
    throw new Error('Missing required columns: ' + missing.join(', '));
  }

  const selectedYear = normalizeSelection(options && options.year);
  const selectedSeason = normalizeSelection(options && options.season);
  const selectedWeek = normalizeSelection(options && options.week);

  const allRows = values.slice(1).filter(function (row) {
    return row.some(function (cell) {
      return cell !== '' && cell !== null && cell !== undefined;
    });
  });

  const filteredRows = allRows.filter(function (row) {
    const rowYear = normalizeSelection(row[indexes.year]);
    const rowSeason = normalizeSelection(row[indexes.season]);
    const rowWeek = normalizeSelection(row[indexes.week]);

    if (selectedYear && rowYear !== selectedYear) return false;
    if (selectedSeason && rowSeason !== selectedSeason) return false;
    if (selectedWeek && rowWeek !== selectedWeek) return false;

    return true;
  });

  const years = uniqueSorted(allRows.map(function (row) {
    return String(row[indexes.year] || '').trim();
  }));

  const seasonRows = !selectedYear ? allRows : allRows.filter(function (row) {
    return normalizeSelection(row[indexes.year]) === selectedYear;
  });

  const seasons = uniqueSorted(seasonRows.map(function (row) {
    return String(row[indexes.season] || '').trim();
  }));

  const weekRows = !selectedYear ? allRows : allRows.filter(function (row) {
    return normalizeSelection(row[indexes.year]) === selectedYear;
  });

  const filteredWeeks = !selectedSeason ? weekRows : weekRows.filter(function (row) {
    return normalizeSelection(row[indexes.season]) === selectedSeason;
  });

  const weeks = sortWeeks(uniqueSorted(filteredWeeks.map(function (row) {
    return String(row[indexes.week] || '').trim();
  })));

  const weekResults = {};

  const filteredSet = new Set(filteredRows);

  allRows.forEach(function (row) {
    const weekName = String(row[indexes.week] || '').trim();
    const weekValue = [normalizeSelection(row[indexes.year]), normalizeSelection(row[indexes.season]), weekName].join('|');
    const teamGroup = normalizeTeam(row[indexes.teams]);
    const playerName = String(row[indexes.players] || '').trim();
    const goals = Number(row[indexes.goals]) || 0;

    if (!weekName || !playerName) {
      return;
    }

    if (!weekResults[weekValue]) {
      weekResults[weekValue] = {
        bibsGoals: 0,
        nonBibsGoals: 0,
        players: {}
      };
    }

    applyScoreOverride(weekResults[weekValue], row, indexes);
    if (filteredSet.has(row)) weekResults[weekValue].inFilter = true;

    if (teamGroup === 'bibs') {
      weekResults[weekValue].bibsGoals += goals;
    } else if (teamGroup === 'nonbibs') {
      weekResults[weekValue].nonBibsGoals += goals;
    }

    if (!weekResults[weekValue].players[playerName]) {
      weekResults[weekValue].players[playerName] = {
        player: playerName,
        side: teamGroup,
        goals: 0
      };
    }

    weekResults[weekValue].players[playerName].goals += goals;
  });

  const playerStats = {};

  const gameKeys = Object.keys(weekResults).sort(function (a, b) {
    const partsA = a.split('|');
    const partsB = b.split('|');
    if (partsA[0] !== partsB[0]) return partsA[0].localeCompare(partsB[0]);
    return weekSortValue(partsA[2]) - weekSortValue(partsB[2]);
  });
  let lastFilteredIndex = -1;
  gameKeys.forEach(function (key, index) {
    if (weekResults[key].inFilter) lastFilteredIndex = index;
  });
  const formGames = lastFilteredIndex === -1 ? [] : gameKeys.slice(Math.max(0, lastFilteredIndex - 4), lastFilteredIndex + 1).reverse();
  const formResults = {};

  gameKeys.forEach(function (weekValue) {
    const currentWeek = weekResults[weekValue];
    const playerNames = Object.keys(currentWeek.players);

    if (currentWeek.bibsOverride !== undefined) currentWeek.bibsGoals = currentWeek.bibsOverride;
    if (currentWeek.nonBibsOverride !== undefined) currentWeek.nonBibsGoals = currentWeek.nonBibsOverride;

    let winner = 'draw';
    if (currentWeek.bibsGoals > currentWeek.nonBibsGoals) {
      winner = 'bibs';
    } else if (currentWeek.nonBibsGoals > currentWeek.bibsGoals) {
      winner = 'nonbibs';
    }

    playerNames.forEach(function (playerName) {
      const player = currentWeek.players[playerName];
      const side = player.side;

      formResults[playerName + '|' + weekValue] = winner === 'draw' ? 'D' : (winner === side ? 'W' : 'L');

      if (!currentWeek.inFilter) return;

      if (!playerStats[playerName]) {
        playerStats[playerName] = {
          player: playerName,
          played: 0,
          wins: 0,
          draws: 0,
          losses: 0,
          points: 0,
          goals: 0
        };
      }

      playerStats[playerName].played += 1;
      playerStats[playerName].goals += player.goals;

      if (winner === 'draw') {
        playerStats[playerName].draws += 1;
        playerStats[playerName].points += 1;
      } else if (winner === 'bibs' && side === 'bibs') {
        playerStats[playerName].wins += 1;
        playerStats[playerName].points += 3;
      } else if (winner === 'nonbibs' && side === 'nonbibs') {
        playerStats[playerName].wins += 1;
        playerStats[playerName].points += 3;
      } else {
        playerStats[playerName].losses += 1;
      }
    });
  });

  const leagueRows = Object.keys(playerStats)
    .map(function (playerName) {
      return {
        player: playerName,
        played: playerStats[playerName].played,
        wins: playerStats[playerName].wins,
        draws: playerStats[playerName].draws,
        losses: playerStats[playerName].losses,
        points: playerStats[playerName].points,
        ppg: playerStats[playerName].played ? playerStats[playerName].points / playerStats[playerName].played : 0,
        goals: playerStats[playerName].goals,
        form: formGames.map(function (gameKey) {
          return formResults[playerName + '|' + gameKey] || '-';
        })
      };
    })
    .sort(function (a, b) {
      if (b.points !== a.points) return b.points - a.points;
      if (b.goals !== a.goals) return b.goals - a.goals;
      if (b.wins !== a.wins) return b.wins - a.wins;
      return a.player.localeCompare(b.player);
    })
    .map(function (row, index) {
      row.position = index + 1;
      return row;
    });

  return {
    rows: leagueRows,
    filters: {
      years: years,
      seasons: seasons,
      weeks: weeks
    },
    selected: {
      year: selectedYear || '',
      season: selectedSeason || '',
      week: selectedWeek || ''
    }
  };
}

function normalizeSelection(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function normalizeTeam(value) {
  const text = String(value || '').trim().toLowerCase();
  if (!text) return '';
  const compact = text.replace(/[^a-z]/g, '');
  if (compact === 'bibs') return 'bibs';
  if (compact === 'nonbibs' || compact === 'nonbibs') return 'nonbibs';
  if (compact.indexOf('non') !== -1 && compact.indexOf('bibs') !== -1) return 'nonbibs';
  return compact;
}

function uniqueSorted(values) {
  return Array.from(new Set(values.filter(function (value) {
    return value !== '' && value !== null && value !== undefined;
  }).map(function (value) {
    return String(value).trim();
  }))).sort();
}

function findScoreColumn(headers, key) {
  return headers.findIndex(function (header) {
    return header.replace(/[^a-z]/g, '') === key;
  });
}

function readScore(value) {
  if (value === '' || value === null || value === undefined) return undefined;
  const number = Number(value);
  return isNaN(number) ? undefined : number;
}

// A typed score on any row of the game overrides the total of player goals.
function applyScoreOverride(game, row, indexes) {
  const bibs = indexes.bibsScore === -1 ? undefined : readScore(row[indexes.bibsScore]);
  const nonBibs = indexes.nonBibsScore === -1 ? undefined : readScore(row[indexes.nonBibsScore]);
  if (bibs !== undefined) game.bibsOverride = bibs;
  if (nonBibs !== undefined) game.nonBibsOverride = nonBibs;
}

function getGamesData(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Full Raw data');
  if (!sheet) throw new Error("Sheet 'Full Raw data' not found.");

  const values = sheet.getDataRange().getValues();
  const empty = { games: [], filters: { years: [], seasons: [], weeks: [] }, selected: { year: '', season: '', week: '' } };
  if (!values || values.length < 2) return empty;

  const headers = values[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const indexes = {
    year: headers.indexOf('year'),
    season: headers.indexOf('season'),
    week: headers.indexOf('week'),
    teams: headers.indexOf('teams'),
    players: headers.indexOf('players'),
    goals: headers.indexOf('goals'),
    bibsScore: findScoreColumn(headers, 'bibsscore'),
    nonBibsScore: findScoreColumn(headers, 'nonbibsscore')
  };
  ['year', 'season', 'week', 'teams', 'players', 'goals'].forEach(function (key) {
    if (indexes[key] === -1) throw new Error('Missing required column: ' + key);
  });

  const selYear = normalizeSelection(options && options.year);
  const selSeason = normalizeSelection(options && options.season);
  const selWeek = normalizeSelection(options && options.week);

  const rows = values.slice(1).filter(function (row) {
    return row.some(function (c) { return c !== '' && c !== null && c !== undefined; });
  });

  const years = uniqueSorted(rows.map(function (r) { return normalizeSelection(r[indexes.year]); }));
  const seasons = uniqueSorted(rows.filter(function (r) {
    return !selYear || normalizeSelection(r[indexes.year]) === selYear;
  }).map(function (r) { return normalizeSelection(r[indexes.season]); }));
  const weeks = sortWeeks(uniqueSorted(rows.filter(function (r) {
    return (!selYear || normalizeSelection(r[indexes.year]) === selYear) &&
      (!selSeason || normalizeSelection(r[indexes.season]) === selSeason);
  }).map(function (r) { return normalizeSelection(r[indexes.week]); })));

  const games = {};
  const order = [];

  rows.forEach(function (row) {
    const year = normalizeSelection(row[indexes.year]);
    const season = normalizeSelection(row[indexes.season]);
    const week = normalizeSelection(row[indexes.week]);
    if (!week) return;
    if (selYear && year !== selYear) return;
    if (selSeason && season !== selSeason) return;
    if (selWeek && week !== selWeek) return;

    const key = [year, season, week].join('|');
    if (!games[key]) {
      games[key] = { year: year, season: season, week: week, bibs: [], nonBibs: [], bibsGoals: 0, nonBibsGoals: 0 };
      order.push(key);
    }
    const game = games[key];
    applyScoreOverride(game, row, indexes);

    const side = normalizeTeam(row[indexes.teams]);
    const player = String(row[indexes.players] || '').trim();
    const goals = Number(row[indexes.goals]) || 0;
    if (side === 'bibs') game.bibsGoals += goals;
    else if (side === 'nonbibs') game.nonBibsGoals += goals;
    if (!player) return;
    if (side === 'bibs') game.bibs.push({ player: player, goals: goals });
    else if (side === 'nonbibs') game.nonBibs.push({ player: player, goals: goals });
  });

  const result = order.map(function (key) {
    const g = games[key];
    const bibsScore = g.bibsOverride !== undefined ? g.bibsOverride : g.bibsGoals;
    const nonBibsScore = g.nonBibsOverride !== undefined ? g.nonBibsOverride : g.nonBibsGoals;
    return {
      year: g.year, season: g.season, week: g.week,
      bibs: g.bibs, nonBibs: g.nonBibs,
      bibsScore: bibsScore, nonBibsScore: nonBibsScore,
      estimated: g.bibsOverride !== undefined || g.nonBibsOverride !== undefined
    };
  }).sort(function (a, b) {
    if (a.year !== b.year) return b.year.localeCompare(a.year);
    return weekSortValue(b.week) - weekSortValue(a.week);
  });

  return {
    games: result,
    filters: { years: years, seasons: seasons, weeks: weeks },
    selected: { year: selYear, season: selSeason, week: selWeek }
  };
}

var MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

// Turns text like "6th September" into a calendar-order number (month * 100 + day).
function weekSortValue(week) {
  const text = String(week || '').toLowerCase();
  const monthIndex = MONTHS.findIndex(function (m) { return text.indexOf(m) !== -1; });
  const dayMatch = text.match(/\d+/);
  if (monthIndex === -1) return Number.MAX_SAFE_INTEGER;
  return (monthIndex + 1) * 100 + (dayMatch ? Number(dayMatch[0]) : 0);
}

function sortWeeks(weeks) {
  return weeks.slice().sort(function (a, b) {
    return weekSortValue(a) - weekSortValue(b) || a.localeCompare(b);
  });
}

// ---- Admin: add a week's game to the raw data sheet ----

function getRawSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Full Raw data');
  if (!sheet) throw new Error("Sheet 'Full Raw data' not found.");
  return sheet;
}

function getAdminOptions() {
  const values = getRawSheet().getDataRange().getValues();
  const headers = values[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const col = function (name) { return headers.indexOf(name); };
  const unique = function (index) {
    if (index === -1) return [];
    return uniqueSorted(values.slice(1).map(function (r) { return r[index]; }));
  };
  return {
    years: unique(col('year')),
    seasons: unique(col('season')),
    players: unique(col('players'))
  };
}

// The edit code is kept in Project Settings > Script properties as ADMIN_CODE.
function checkAdminCode(code) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_CODE');
  if (!expected) throw new Error('No ADMIN_CODE script property has been set.');
  if (String(code || '') !== expected) throw new Error('Incorrect edit code.');
}

function addGame(code, game) {
  checkAdminCode(code);

  const year = normalizeSelection(game && game.year);
  const season = normalizeSelection(game && game.season);
  const week = normalizeSelection(game && game.week);
  if (!year || !season || !week) throw new Error('Year, season and week are required.');

  const teams = [
    { label: 'Bibs', number: 1, players: (game.bibs || []) },
    { label: 'Non bibs', number: 2, players: (game.nonBibs || []) }
  ];

  const bibsScore = readScore(game.bibsScore);
  const nonBibsScore = readScore(game.nonBibsScore);
  [game.bibsScore, game.nonBibsScore].forEach(function (value) {
    if (value !== '' && value !== null && value !== undefined && (readScore(value) === undefined || readScore(value) < 0)) {
      throw new Error('Scores must be numbers of 0 or more.');
    }
  });

  const newRows = [];
  teams.forEach(function (team) {
    team.players.forEach(function (p) {
      const name = String(p && p.name || '').trim();
      if (!name) return;
      const goals = p.goals === '' || p.goals === null || p.goals === undefined ? 0 : Number(p.goals);
      if (isNaN(goals) || goals < 0 || Math.floor(goals) !== goals) {
        throw new Error('Goals for ' + name + ' must be a whole number of 0 or more.');
      }
      newRows.push({ team: team, name: name, goals: goals });
    });
  });
  if (!newRows.length) throw new Error('Add at least one player.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getRawSheet();
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
    const idx = {
      year: headers.indexOf('year'),
      season: headers.indexOf('season'),
      week: headers.indexOf('week'),
      teams: headers.indexOf('teams'),
      team: headers.indexOf('team'),
      players: headers.indexOf('players'),
      goals: headers.indexOf('goals'),
      bibsScore: findScoreColumn(headers, 'bibsscore'),
      nonBibsScore: findScoreColumn(headers, 'nonbibsscore')
    };
    ['year', 'season', 'week', 'teams', 'team', 'players', 'goals'].forEach(function (key) {
      if (idx[key] === -1) throw new Error('Missing column: ' + key);
    });

    const exists = values.slice(1).some(function (r) {
      return normalizeSelection(r[idx.year]) === year &&
        normalizeSelection(r[idx.season]) === season &&
        normalizeSelection(r[idx.week]).toLowerCase() === week.toLowerCase();
    });
    if (exists) throw new Error('A game for ' + year + ' / ' + season + ' / ' + week + ' already exists.');

    const output = newRows.map(function (item, i) {
      const row = new Array(headers.length).fill('');
      row[idx.year] = isNaN(Number(year)) ? year : Number(year);
      row[idx.season] = season;
      row[idx.week] = week;
      row[idx.teams] = item.team.label;
      row[idx.team] = item.team.number;
      row[idx.players] = item.name;
      row[idx.goals] = item.goals;
      if (i === 0) {
        if (idx.bibsScore !== -1 && bibsScore !== undefined) row[idx.bibsScore] = bibsScore;
        if (idx.nonBibsScore !== -1 && nonBibsScore !== undefined) row[idx.nonBibsScore] = nonBibsScore;
      }
      return row;
    });

    if ((bibsScore !== undefined && idx.bibsScore === -1) || (nonBibsScore !== undefined && idx.nonBibsScore === -1)) {
      throw new Error("Add 'Bibs Score' and 'Non-Bibs Score' headers to the sheet first.");
    }

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, idx.week + 1, output.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, output.length, headers.length).setValues(output);
    return { added: output.length };
  } finally {
    lock.releaseLock();
  }
}

// Year and season of the most recent game in the sheet.
function getLatestSeason() {
  const values = getRawSheet().getDataRange().getValues();
  const headers = values[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const yearIdx = headers.indexOf('year');
  const seasonIdx = headers.indexOf('season');
  const weekIdx = headers.indexOf('week');
  let best = null;

  values.slice(1).forEach(function (row) {
    const year = normalizeSelection(row[yearIdx]);
    const season = normalizeSelection(row[seasonIdx]);
    if (!year || !season) return;
    const score = Number(year) * 10000 + weekSortValue(row[weekIdx]);
    if (!best || score > best.score) best = { year: year, season: season, score: score };
  });

  return best ? { year: best.year, season: best.season } : { year: '', season: '' };
}
