/*****************************************************************************
 * GOOGLE ADS ACCOUNT GUARDRAILS
 * Version 2.0
 *
 * Checks one Google Ads account every day and sends an email when something
 * looks wrong: spend too high or too low, no spend at all, conversions or
 * conversion value dropping, ROAS / CPA off target, ads disapproved, etc.
 * If the script itself fails, you get an email about that as well.
 *
 * HOW TO INSTALL
 *   1. In Google Ads open: Tools > Bulk actions > Scripts > (+) New script.
 *   2. Delete everything in the editor and paste this whole file.
 *   3. Change the values in the SETUP section below. Nothing else needs editing.
 *   4. Click Authorize, then Preview. Check the Logs tab for the results.
 *   5. Save, then set Frequency to "Daily", preferably after 6:00 so that
 *      yesterday's data is complete.
 *
 * EDITING RULES (to avoid breaking the script)
 *   - Only change what is after the "=" sign and before the ";".
 *   - Text goes between quotes:      'HUF'
 *   - Numbers have no quotes or spaces: 150000   (not '150 000')
 *   - On/off switches are lowercase:   true  or  false
 *   - Lines starting with // are explanations and are ignored by Google Ads.
 *****************************************************************************/


// ================================ SETUP =====================================

// ---------- 1. EMAIL AND ACCOUNT -----------------------------------------------

// Who receives the alerts. Several addresses: separate them with a comma.
// Example: 'anna@hdmarketing.hu, peter@hdmarketing.hu'
var EMAIL_RECIPIENTS = 'name@example.com';

// Optional friendly name shown in the email subject.
// Leave it empty ('') to use the account name from Google Ads.
var ACCOUNT_LABEL = '';

// Currency of this account: 'HUF', 'EUR' or 'USD'.
// Every money amount in this setup is understood in this currency.
// If it does not match the real account currency, the email will warn you.
var CURRENCY = 'HUF';


// ---------- 2. WHICH DAYS ARE CHECKED AND WHAT IS "NORMAL" ----------------------

// How many recent days are checked. The period always ends YESTERDAY,
// because today is not finished yet.
//   1 = yesterday only. Fastest reaction, but one weak day can cause alerts.
//   3 = the last 3 days. Good balance, recommended.
//   7 = the last 7 days. Calm, but problems are noticed later.
// Daily limits below are compared with the DAILY AVERAGE of these days.
// Note: conversions often arrive a few days after the click, so the most
// recent days always look a bit weaker. With 1 day this is most visible.
var LOOKBACK_DAYS = 3;

// The period used as "normal" in every "compared to average" check:
//   'LAST_MONTH'    = the whole previous calendar month
//                     (e.g. all of September when the script runs in October).
//   'PREVIOUS_DAYS' = the days right before the checked period
//                     (how many: see BASELINE_DAYS). Follows seasons better.
var BASELINE_PERIOD = 'LAST_MONTH';

// Only used with 'PREVIOUS_DAYS': how many days make up the "normal" period.
var BASELINE_DAYS = 30;

// How the normal daily value is calculated from that period:
//   'MEAN'   = simple average: total divided by the number of days.
//   'MEDIAN' = the middle day when all days are sorted. One-off extreme days
//              (a sale, an outage) do not distort it, so it is more stable.
var BASELINE_AVERAGE = 'MEAN';

// Safety net for small accounts. Comparisons of conversions, conversion value,
// ROAS and CPA are skipped when fewer conversions than this would normally
// happen in the checked days, because tiny numbers jump around randomly.
// Example: normally 1 conversion a day and LOOKBACK_DAYS = 3 means 3 expected,
// which is below 5, so those comparisons are skipped that day.
var MIN_EXPECTED_CONVERSIONS = 5;

// true  = also send a short "all clear" email when nothing is wrong.
// false = only send an email when something needs attention.
var SEND_ALL_CLEAR_EMAIL = false;


// ---------- 3. SPEND -----------------------------------------------------------

// MINIMUM DAILY SPEND
// Alert if the average daily spend of the checked days is below this amount.
var CHECK_SPEND_MIN = true;
var SPEND_MIN_DAILY = 5000;

// MAXIMUM DAILY SPEND
// Alert if the average daily spend of the checked days is above this amount.
var CHECK_SPEND_MAX = true;
var SPEND_MAX_DAILY = 100000;

// SPEND COMPARED TO NORMAL
// Compares the average daily spend of the checked days with the normal daily
// spend (see section 2).
// SPEND_MAX_INCREASE_PERCENT = 30 means: alert if spend is more than 30% above
// normal (overspend). SPEND_MAX_DECREASE_PERCENT = 50 means: alert if spend is
// more than 50% below normal (underspend). Set either one to 0 to ignore that side.
var CHECK_SPEND_VS_AVERAGE = true;
var SPEND_MAX_INCREASE_PERCENT = 30;
var SPEND_MAX_DECREASE_PERCENT = 50;

// MONTHLY BUDGET OF THE WHOLE ACCOUNT
// Alert if this month's total spend (including today so far) is above
// MONTHLY_BUDGET. With MONTHLY_BUDGET_FORECAST = true you are also warned
// early, when the current pace would end the month above the budget.
var CHECK_MONTHLY_BUDGET = true;
var MONTHLY_BUDGET = 3000000;
var MONTHLY_BUDGET_FORECAST = true;

// MONTHLY LIMIT PER CAMPAIGN
// Alert for every active campaign that spent more than this amount this month.
var CHECK_CAMPAIGN_MONTHLY_BUDGET = false;
var CAMPAIGN_MONTHLY_BUDGET = 300000;

// SEARCH IMPRESSION SHARE LOST TO BUDGET
// Alert for enabled Search campaigns that missed more than this percent of the
// impressions they could have had in the checked days, because their daily
// budget ran out. Example: 20 means "lost more than 20% to budget".
// Campaigns that spent less than LOST_IS_MIN_SPEND in the checked days are
// ignored, so tiny test campaigns do not cause alerts.
// A budget-limited campaign is not always a problem: if it is profitable, it is
// losing sales; if its budget is capped on purpose, switch this check off.
var CHECK_LOST_IS_BUDGET = true;
var LOST_IS_BUDGET_MAX_PERCENT = 20;
var LOST_IS_MIN_SPEND = 10000;


// ---------- 4. NO SPEND, BALANCE, STOPPED CAMPAIGNS ----------------------------

// NO SPEND YESTERDAY
// Alert if the account spent nothing at all yesterday. The usual reasons are:
// prepaid balance ran out, failed payment, suspended account, or every
// campaign paused. Google does not let scripts read a prepaid balance, so this
// is the most reliable "out of money" signal.
var CHECK_NO_SPEND = true;

// ACCOUNT BUDGET (only for accounts billed by monthly invoice / budget orders)
// Alert if the approved account budget will run out within this many days at
// the current spending pace, or if it ends within this many days.
// Accounts without such a budget are simply skipped, this causes no error.
var CHECK_ACCOUNT_BUDGET = true;
var ACCOUNT_BUDGET_MIN_DAYS_LEFT = 7;

// STOPPED CAMPAIGNS
// Alert for active campaigns that were showing ads in the normal period but had
// zero impressions yesterday (e.g. budget, policy or feed problem, or a
// schedule mistake). Skipped when the whole account had no spend.
var CHECK_STALLED_CAMPAIGNS = true;


// ---------- 5. CONVERSIONS ------------------------------------------------------
// Only conversions counted in the "Conversions" column (primary conversion
// actions) are used. Secondary actions are ignored.

// MINIMUM CONVERSIONS
// Alert if the TOTAL number of conversions in the checked days is below this.
// Example: LOOKBACK_DAYS = 3 and CONVERSIONS_MIN = 3 means "fewer than 3
// conversions in the last 3 days".
var CHECK_CONVERSIONS_MIN = true;
var CONVERSIONS_MIN = 3;

// CONVERSIONS COMPARED TO NORMAL
// Alert if conversions are more than this percent below the normal amount.
var CHECK_CONVERSIONS_VS_AVERAGE = true;
var CONVERSIONS_MAX_DROP_PERCENT = 50;

// EACH PRIMARY CONVERSION ACTION SEPARATELY
// Alert if a primary conversion action (e.g. "Purchase", "Lead form",
// "Phone call") that normally records conversions has recorded ZERO in the
// last CONVERSION_ACTIONS_LOOKBACK_DAYS days. This catches one broken tag even
// when the other conversions still work. Actions that rarely convert are
// skipped automatically (see MIN_EXPECTED_CONVERSIONS).
// To exclude actions, list their exact names, e.g. ['Phone call', 'Chat start'].
var CHECK_CONVERSION_ACTIONS = true;
var CONVERSION_ACTIONS_LOOKBACK_DAYS = 7;
var CONVERSION_ACTIONS_IGNORE = [];


// ---------- 6. CONVERSION VALUE ------------------------------------------------
// Only relevant if conversions have a value (typically webshops).

// MINIMUM CONVERSION VALUE
// Alert if the total conversion value of the checked days is below this amount.
var CHECK_CONV_VALUE_MIN = false;
var CONV_VALUE_MIN = 100000;

// CONVERSION VALUE COMPARED TO NORMAL
// Alert if conversion value is more than this percent below the normal amount.
var CHECK_CONV_VALUE_VS_AVERAGE = true;
var CONV_VALUE_MAX_DROP_PERCENT = 50;


// ---------- 7. EFFICIENCY: ROAS AND CPA ----------------------------------------

// ROAS (conversion value divided by cost)
// ROAS_TARGET_PERCENT is written the same way as Target ROAS in Google Ads:
// 400 means every 1 HUF/EUR/USD spent brings 4 back in conversion value.
// ROAS_COMPARE_TO decides what the ROAS of the checked days is compared with:
//   'TARGET'  = the ROAS_TARGET_PERCENT number above.
//   'AVERAGE' = the ROAS of the normal period (see section 2).
// ROAS_MAX_DEVIATION_PERCENT = 50 means: alert if ROAS is more than 50% away
// from that reference (with target 400: below 200 or above 600).
// ROAS_DIRECTION: 'BELOW' = only alert when ROAS is too low,
//                 'ABOVE' = only when too high,
//                 'BOTH'  = both. A suspiciously high ROAS often means a
//                           conversion or value is being counted twice.
// Skipped automatically if the account tracks no conversion value.
var CHECK_ROAS = true;
var ROAS_TARGET_PERCENT = 400;
var ROAS_COMPARE_TO = 'TARGET';
var ROAS_MAX_DEVIATION_PERCENT = 50;
var ROAS_DIRECTION = 'BOTH';

// CPA (cost per conversion) COMPARED TO NORMAL
// Useful for lead generation accounts where ROAS means nothing.
// Alert if the cost of one conversion is more than this percent above normal.
var CHECK_CPA = false;
var CPA_MAX_INCREASE_PERCENT = 50;


// ---------- 8. ADS ---------------------------------------------------------------

// DISAPPROVED ADS
// Alert if an active ad in an active ad group and campaign is disapproved by
// Google policy. Repeats every day until the ad is fixed or paused.
var CHECK_DISAPPROVED_ADS = true;


// ---------- 9. CHANGE HISTORY ----------------------------------------------------
// Reads the account's change history and alerts on large changes, whoever made
// them: a colleague, the client, an automated rule, another script, or a Google
// recommendation that was applied automatically. The email shows who did it.

// BIG BUDGET CHANGES
// Alert if a campaign budget was raised or lowered by at least this percent in
// a single change. Example: 30 means 100 000 -> 130 000 or 100 000 -> 70 000.
var CHECK_BUDGET_CHANGES = true;
var BUDGET_CHANGE_PERCENT = 30;

// BIG BID CHANGES
// Alert if a bid or bidding target changed by at least this percent in a single
// change. Covered: Target ROAS, Target CPA, ad group and keyword max. CPC bids,
// and portfolio bid strategies. Switching a campaign to a different bid strategy
// (e.g. from Maximize clicks to Target ROAS) is always reported.
var CHECK_BID_CHANGES = true;
var BID_CHANGE_PERCENT = 30;

// How many days of history to read. The period ends yesterday.
// 1 = yesterday only. Recommended for a daily schedule: every change is
//     reported exactly once, the morning after it was made.
// Google keeps at most 30 days, so the maximum here is 29.
var CHANGE_HISTORY_DAYS = 1;

// Changes made by these people are not reported, e.g. your own team, so you
// only hear about changes made by others or by automation.
// Exact email addresses, e.g. ['anna@hdmarketing.hu', 'peter@hdmarketing.hu']
// Leave it as [] to report everyone's changes.
var CHANGE_HISTORY_IGNORE_USERS = [];

// ============================== END OF SETUP ================================
// Do not edit below this line.


var DAY_MS = 86400000;
var MAX_LIST_ITEMS = 25;

function main() {
  var ctx = { alerts: [], errors: [] };
  try {
    setup(ctx);
    var checks = [
      [CHECK_NO_SPEND, 'No spend yesterday', checkNoSpend],
      [CHECK_SPEND_MIN, 'Minimum daily spend', checkSpendMin],
      [CHECK_SPEND_MAX, 'Maximum daily spend', checkSpendMax],
      [CHECK_SPEND_VS_AVERAGE, 'Spend compared to normal', checkSpendVsAverage],
      [CHECK_MONTHLY_BUDGET, 'Monthly account budget', checkMonthlyBudget],
      [CHECK_CAMPAIGN_MONTHLY_BUDGET, 'Monthly campaign limit', checkCampaignMonthlyBudget],
      [CHECK_LOST_IS_BUDGET, 'Impression share lost to budget', checkLostIsBudget],
      [CHECK_ACCOUNT_BUDGET, 'Account budget', checkAccountBudget],
      [CHECK_STALLED_CAMPAIGNS, 'Stopped campaigns', checkStalledCampaigns],
      [CHECK_CONVERSIONS_MIN, 'Minimum conversions', checkConversionsMin],
      [CHECK_CONVERSIONS_VS_AVERAGE, 'Conversions compared to normal', checkConversionsVsAverage],
      [CHECK_CONVERSION_ACTIONS, 'Primary conversion actions', checkConversionActions],
      [CHECK_CONV_VALUE_MIN, 'Minimum conversion value', checkConvValueMin],
      [CHECK_CONV_VALUE_VS_AVERAGE, 'Conversion value compared to normal', checkConvValueVsAverage],
      [CHECK_ROAS, 'ROAS', checkRoas],
      [CHECK_CPA, 'CPA compared to normal', checkCpa],
      [CHECK_DISAPPROVED_ADS, 'Disapproved ads', checkDisapprovedAds],
      [CHECK_BUDGET_CHANGES, 'Budget changes', checkBudgetChanges],
      [CHECK_BID_CHANGES, 'Bid changes', checkBidChanges]
    ];
    checks.forEach(function (c) {
      if (c[0] !== true) return;
      try {
        c[2](ctx);
      } catch (e) {
        recordError(ctx, c[1], e);
      }
    });
  } catch (e) {
    recordError(ctx, 'Script setup', e);
  }
  sendReport(ctx);
  if (ctx.errors.length) {
    throw new Error(ctx.errors.length + ' guardrail check(s) failed, see the email or the logs above.');
  }
}

// ---------- Setup and data ----------

function setup(ctx) {
  var account = AdsApp.currentAccount();
  ctx.account = {
    name: ACCOUNT_LABEL || account.getName(),
    id: account.getCustomerId(),
    currency: account.getCurrencyCode(),
    tz: account.getTimeZone()
  };
  validateSettings();

  ctx.today = parseYmd(Utilities.formatDate(new Date(), ctx.account.tz, 'yyyy-MM-dd'));
  ctx.yesterday = addDays(ctx.today, -1);
  ctx.window = period(addDays(ctx.today, -LOOKBACK_DAYS), ctx.yesterday);
  ctx.baseline = baselinePeriod(ctx.today, ctx.window.start);

  ctx.w = loadDaily(ctx.window);
  ctx.b = loadDaily(ctx.baseline);
  ctx.y = LOOKBACK_DAYS === 1 ? ctx.w : loadDaily(period(ctx.yesterday, ctx.yesterday));

  Logger.log('Account: ' + ctx.account.name + ' (' + ctx.account.id + ')');
  Logger.log('Checked period: ' + periodLabel(ctx.window) + ', normal period: ' + periodLabel(ctx.baseline));

  if (ctx.account.currency !== CURRENCY) {
    addAlert(ctx, {
      severity: 'warning',
      title: 'Currency setting does not match the account',
      message: 'The script is set to ' + CURRENCY + ' but the account uses ' + ctx.account.currency +
        '. Every money limit is probably wrong by a large factor.',
      hint: 'Set CURRENCY to \'' + ctx.account.currency + '\' and review all money amounts in the setup.'
    });
  }
}

function validateSettings() {
  var problems = [];
  if (!/@/.test(EMAIL_RECIPIENTS) || /example\.com/.test(EMAIL_RECIPIENTS)) problems.push('EMAIL_RECIPIENTS is not set.');
  if (!(LOOKBACK_DAYS >= 1 && LOOKBACK_DAYS <= 90 && LOOKBACK_DAYS % 1 === 0)) problems.push('LOOKBACK_DAYS must be a whole number between 1 and 90.');
  if (['LAST_MONTH', 'PREVIOUS_DAYS'].indexOf(BASELINE_PERIOD) < 0) problems.push('BASELINE_PERIOD must be \'LAST_MONTH\' or \'PREVIOUS_DAYS\'.');
  if (!(BASELINE_DAYS >= 1 && BASELINE_DAYS <= 365 && BASELINE_DAYS % 1 === 0)) problems.push('BASELINE_DAYS must be a whole number between 1 and 365.');
  if (['MEAN', 'MEDIAN'].indexOf(BASELINE_AVERAGE) < 0) problems.push('BASELINE_AVERAGE must be \'MEAN\' or \'MEDIAN\'.');
  if (['TARGET', 'AVERAGE'].indexOf(ROAS_COMPARE_TO) < 0) problems.push('ROAS_COMPARE_TO must be \'TARGET\' or \'AVERAGE\'.');
  if (['BOTH', 'BELOW', 'ABOVE'].indexOf(ROAS_DIRECTION) < 0) problems.push('ROAS_DIRECTION must be \'BOTH\', \'BELOW\' or \'ABOVE\'.');
  if (!(CONVERSION_ACTIONS_LOOKBACK_DAYS >= 1 && CONVERSION_ACTIONS_LOOKBACK_DAYS % 1 === 0)) problems.push('CONVERSION_ACTIONS_LOOKBACK_DAYS must be a whole number of at least 1.');
  if (!(CHANGE_HISTORY_DAYS >= 1 && CHANGE_HISTORY_DAYS <= 29 && CHANGE_HISTORY_DAYS % 1 === 0)) problems.push('CHANGE_HISTORY_DAYS must be a whole number between 1 and 29.');
  if (problems.length) throw new Error('Setup problem: ' + problems.join(' '));
}

function baselinePeriod(today, windowStart) {
  if (BASELINE_PERIOD === 'LAST_MONTH') {
    var y = today.getUTCFullYear(), m = today.getUTCMonth();
    return period(new Date(Date.UTC(y, m - 1, 1, 12)), new Date(Date.UTC(y, m, 0, 12)));
  }
  return period(addDays(windowStart, -BASELINE_DAYS), addDays(windowStart, -1));
}

function loadDaily(p) {
  var byDate = {};
  var rows = AdsApp.search(
    'SELECT segments.date, metrics.cost_micros, metrics.conversions, metrics.conversions_value, metrics.impressions ' +
    'FROM customer WHERE segments.date BETWEEN \'' + p.from + '\' AND \'' + p.to + '\'');
  while (rows.hasNext()) {
    var r = rows.next();
    byDate[r.segments.date] = {
      cost: Number(r.metrics.costMicros || 0) / 1e6,
      conv: Number(r.metrics.conversions || 0),
      value: Number(r.metrics.conversionsValue || 0),
      impr: Number(r.metrics.impressions || 0)
    };
  }
  var stats = { period: p, days: [], n: 0, cost: 0, conv: 0, value: 0, impr: 0 };
  for (var d = p.start; d <= p.end; d = addDays(d, 1)) {
    var day = byDate[ymd(d)] || { cost: 0, conv: 0, value: 0, impr: 0 };
    stats.days.push(day);
    stats.n++;
    stats.cost += day.cost;
    stats.conv += day.conv;
    stats.value += day.value;
    stats.impr += day.impr;
  }
  return stats;
}

function normalDaily(stats, field) {
  if (BASELINE_AVERAGE === 'MEDIAN') {
    var v = stats.days.map(function (d) { return d[field]; }).sort(function (a, b) { return a - b; });
    var mid = Math.floor(v.length / 2);
    return v.length % 2 ? v[mid] : (v[mid - 1] + v[mid]) / 2;
  }
  return stats[field] / stats.n;
}

function expectedConversions(ctx) {
  return normalDaily(ctx.b, 'conv') * ctx.w.n;
}

function enoughVolume(ctx, checkName) {
  var expected = expectedConversions(ctx);
  if (expected >= MIN_EXPECTED_CONVERSIONS) return true;
  Logger.log(checkName + ': skipped, only ' + num(expected, 1) + ' conversions expected (minimum ' + MIN_EXPECTED_CONVERSIONS + ').');
  return false;
}

function sumCost(from, to) {
  var rows = AdsApp.search('SELECT metrics.cost_micros FROM customer WHERE segments.date BETWEEN \'' + from + '\' AND \'' + to + '\'');
  var total = 0;
  while (rows.hasNext()) total += Number(rows.next().metrics.costMicros || 0) / 1e6;
  return total;
}

// ---------- Checks ----------

function checkNoSpend(ctx) {
  if (ctx.y.cost > 0) return;
  var enabled = 0;
  var rows = AdsApp.search('SELECT campaign.id FROM campaign WHERE campaign.status = \'ENABLED\'');
  while (rows.hasNext()) { rows.next(); enabled++; }
  ctx.noSpendYesterday = true;
  addAlert(ctx, {
    severity: 'critical',
    title: 'No spend yesterday',
    message: 'The account spent nothing on ' + ymd(ctx.yesterday) + '.',
    facts: [['Spend yesterday', money(0)], ['Enabled campaigns', enabled], ['Normal daily spend', money(normalDaily(ctx.b, 'cost'))]],
    hint: enabled ? 'Check Billing (prepaid balance, failed payment) and the account status. Ads are probably not running.'
      : 'No campaign is enabled. If this is intentional you can ignore this alert or switch CHECK_NO_SPEND off.'
  });
}

function checkSpendMin(ctx) {
  var avg = ctx.w.cost / ctx.w.n;
  if (avg >= SPEND_MIN_DAILY) return;
  addAlert(ctx, {
    severity: 'warning',
    title: 'Daily spend below minimum',
    message: 'Average daily spend in ' + periodLabel(ctx.window) + ' was below the set minimum.',
    facts: [['Average daily spend', money(avg)], ['Minimum', money(SPEND_MIN_DAILY)]],
    hint: 'Look for limited budgets, low bids, disapprovals, or a running-low prepaid balance.'
  });
}

function checkSpendMax(ctx) {
  var avg = ctx.w.cost / ctx.w.n;
  if (avg <= SPEND_MAX_DAILY) return;
  addAlert(ctx, {
    severity: 'warning',
    title: 'Daily spend above maximum',
    message: 'Average daily spend in ' + periodLabel(ctx.window) + ' was above the set maximum.',
    facts: [['Average daily spend', money(avg)], ['Maximum', money(SPEND_MAX_DAILY)]],
    hint: 'Check recent budget or bid strategy changes and new campaigns.'
  });
}

function checkSpendVsAverage(ctx) {
  var normal = normalDaily(ctx.b, 'cost');
  if (normal <= 0) {
    Logger.log('Spend compared to normal: skipped, no spend in the normal period.');
    return;
  }
  var avg = ctx.w.cost / ctx.w.n;
  var change = avg / normal - 1;
  var facts = [['Average daily spend', money(avg)], ['Normal daily spend', money(normal)], ['Change', pct(change)]];
  if (SPEND_MAX_INCREASE_PERCENT > 0 && change > SPEND_MAX_INCREASE_PERCENT / 100) {
    addAlert(ctx, {
      severity: 'warning',
      title: 'Overspend compared to normal',
      message: 'Daily spend is more than ' + SPEND_MAX_INCREASE_PERCENT + '% above ' + baselineLabel(ctx) + '.',
      facts: facts,
      hint: 'Check budget changes, bid strategy target changes and new or reactivated campaigns.'
    });
  } else if (SPEND_MAX_DECREASE_PERCENT > 0 && change < -SPEND_MAX_DECREASE_PERCENT / 100) {
    addAlert(ctx, {
      severity: 'warning',
      title: 'Underspend compared to normal',
      message: 'Daily spend is more than ' + SPEND_MAX_DECREASE_PERCENT + '% below ' + baselineLabel(ctx) + '.',
      facts: facts,
      hint: 'Check paused campaigns, disapprovals, limited budgets and the billing balance.'
    });
  }
}

function checkMonthlyBudget(ctx) {
  var monthStart = new Date(Date.UTC(ctx.today.getUTCFullYear(), ctx.today.getUTCMonth(), 1, 12));
  var mtd = sumCost(ymd(monthStart), ymd(ctx.today));
  if (mtd > MONTHLY_BUDGET) {
    addAlert(ctx, {
      severity: 'critical',
      title: 'Monthly budget exceeded',
      message: 'This month\'s spend is already above the monthly budget.',
      facts: [['Spend this month (incl. today)', money(mtd)], ['Monthly budget', money(MONTHLY_BUDGET)], ['Over by', money(mtd - MONTHLY_BUDGET)]],
      hint: 'Pause or limit campaigns for the rest of the month.'
    });
    return;
  }
  var daysDone = ctx.today.getUTCDate() - 1;
  if (!MONTHLY_BUDGET_FORECAST || daysDone < 1) return;
  var daysInMonth = new Date(Date.UTC(ctx.today.getUTCFullYear(), ctx.today.getUTCMonth() + 1, 0, 12)).getUTCDate();
  var doneSpend = sumCost(ymd(monthStart), ymd(ctx.yesterday));
  var forecast = doneSpend / daysDone * daysInMonth;
  if (forecast <= MONTHLY_BUDGET) return;
  addAlert(ctx, {
    severity: 'warning',
    title: 'Monthly budget will be exceeded at the current pace',
    message: 'If spending continues like the first ' + daysDone + ' days, the month ends above budget.',
    facts: [['Spend until yesterday', money(doneSpend)], ['Forecast for the month', money(forecast)], ['Monthly budget', money(MONTHLY_BUDGET)],
      ['Daily spend that fits the budget', money(Math.max(0, MONTHLY_BUDGET - doneSpend) / (daysInMonth - daysDone))]],
    hint: 'Lower daily budgets to the amount above to land on budget.'
  });
}

function checkCampaignMonthlyBudget(ctx) {
  var rows = AdsApp.search(
    'SELECT campaign.name, metrics.cost_micros FROM campaign WHERE segments.date DURING THIS_MONTH ' +
    'AND campaign.status = \'ENABLED\' AND metrics.cost_micros > ' + Math.round(CAMPAIGN_MONTHLY_BUDGET * 1e6) +
    ' ORDER BY metrics.cost_micros DESC');
  var list = [];
  while (rows.hasNext()) {
    var r = rows.next();
    list.push(r.campaign.name + ': ' + money(Number(r.metrics.costMicros) / 1e6));
  }
  if (!list.length) return;
  addAlert(ctx, {
    severity: 'warning',
    title: list.length + ' campaign(s) above the monthly campaign limit',
    message: 'These enabled campaigns spent more than ' + money(CAMPAIGN_MONTHLY_BUDGET) + ' this month.',
    list: list,
    hint: 'Pause or limit them for the rest of the month.'
  });
}

function checkAccountBudget(ctx) {
  var rows = AdsApp.search(
    'SELECT account_budget.name, account_budget.approved_start_date_time, account_budget.approved_end_date_time, ' +
    'account_budget.adjusted_spending_limit_micros, account_budget.adjusted_spending_limit_type, account_budget.amount_served_micros ' +
    'FROM account_budget WHERE account_budget.status = \'APPROVED\'');
  var now = Utilities.formatDate(new Date(), ctx.account.tz, 'yyyy-MM-dd HH:mm:ss');
  var dailySpend = ctx.w.cost / ctx.w.n;
  var found = false;
  while (rows.hasNext()) {
    var b = rows.next().accountBudget;
    if (b.approvedStartDateTime > now || (b.approvedEndDateTime && b.approvedEndDateTime < now)) continue;
    found = true;
    var name = b.name || 'Account budget';
    if (b.adjustedSpendingLimitType !== 'INFINITE' && b.adjustedSpendingLimitMicros != null) {
      var remaining = (Number(b.adjustedSpendingLimitMicros) - Number(b.amountServedMicros || 0)) / 1e6;
      var daysLeft = dailySpend > 0 ? remaining / dailySpend : Infinity;
      if (remaining <= 0 || daysLeft < ACCOUNT_BUDGET_MIN_DAYS_LEFT) {
        addAlert(ctx, {
          severity: remaining <= 0 ? 'critical' : 'warning',
          title: remaining <= 0 ? 'Account budget used up' : 'Account budget running out',
          message: '"' + name + '" has ' + money(Math.max(0, remaining)) + ' left' +
            (isFinite(daysLeft) ? ', about ' + num(Math.max(0, daysLeft), 1) + ' days at the current pace.' : '.'),
          facts: [['Spending limit', money(Number(b.adjustedSpendingLimitMicros) / 1e6)], ['Already spent', money(Number(b.amountServedMicros || 0) / 1e6)],
            ['Average daily spend', money(dailySpend)]],
          hint: 'Ask the client to approve a new or increased budget before ads stop.'
        });
      }
    }
    if (b.approvedEndDateTime) {
      var daysToEnd = (parseYmd(b.approvedEndDateTime.substring(0, 10)) - ctx.today) / DAY_MS;
      if (daysToEnd < ACCOUNT_BUDGET_MIN_DAYS_LEFT) {
        addAlert(ctx, {
          severity: 'warning',
          title: 'Account budget ends soon',
          message: '"' + name + '" ends on ' + b.approvedEndDateTime.substring(0, 10) + '. Ads stop after that date.',
          hint: 'Ask the client to approve a new account budget.'
        });
      }
    }
  }
  if (!found) Logger.log('Account budget: no active budget order, skipped (normal for card or prepaid billing).');
}

function checkStalledCampaigns(ctx) {
  if (ctx.y.cost === 0) {
    Logger.log('Stopped campaigns: skipped, the whole account had no spend yesterday.');
    return;
  }
  var active = {};
  var rows = AdsApp.search(
    'SELECT campaign.id, campaign.name FROM campaign WHERE campaign.status = \'ENABLED\' ' +
    'AND campaign.serving_status NOT IN (\'ENDED\', \'PENDING\') AND metrics.impressions > 0 ' +
    'AND segments.date BETWEEN \'' + ctx.baseline.from + '\' AND \'' + ctx.baseline.to + '\'');
  while (rows.hasNext()) {
    var r = rows.next();
    active[r.campaign.id] = r.campaign.name;
  }
  rows = AdsApp.search(
    'SELECT campaign.id FROM campaign WHERE metrics.impressions > 0 ' +
    'AND segments.date BETWEEN \'' + ymd(ctx.yesterday) + '\' AND \'' + ymd(ctx.yesterday) + '\'');
  while (rows.hasNext()) delete active[rows.next().campaign.id];
  var list = Object.keys(active).map(function (id) { return active[id]; });
  if (!list.length) return;
  addAlert(ctx, {
    severity: 'warning',
    title: list.length + ' enabled campaign(s) had no impressions yesterday',
    message: 'These campaigns are enabled and were showing ads in the normal period, but showed nothing on ' + ymd(ctx.yesterday) + '.',
    list: list,
    hint: 'Check the campaign status column: budget, ad schedule, disapprovals, product feed or bid limits.'
  });
}

function checkConversionsMin(ctx) {
  if (ctx.w.conv >= CONVERSIONS_MIN) return;
  addAlert(ctx, {
    severity: 'warning',
    title: 'Too few conversions',
    message: 'Only ' + num(ctx.w.conv, 1) + ' primary conversions in ' + periodLabel(ctx.window) + '.',
    facts: [['Conversions', num(ctx.w.conv, 1)], ['Minimum', CONVERSIONS_MIN], ['Spend in the same days', money(ctx.w.cost)]],
    hint: 'Check the conversion tags (GTM), the website forms and checkout, and the consent banner.'
  });
}

function checkConversionsVsAverage(ctx) {
  if (!enoughVolume(ctx, 'Conversions compared to normal')) return;
  var expected = expectedConversions(ctx);
  var change = ctx.w.conv / expected - 1;
  if (change >= -CONVERSIONS_MAX_DROP_PERCENT / 100) return;
  addAlert(ctx, {
    severity: 'warning',
    title: 'Conversions dropped compared to normal',
    message: 'Conversions are more than ' + CONVERSIONS_MAX_DROP_PERCENT + '% below ' + baselineLabel(ctx) + '.',
    facts: [['Conversions', num(ctx.w.conv, 1)], ['Normal for ' + ctx.w.n + ' day(s)', num(expected, 1)], ['Change', pct(change)]],
    hint: 'If spend is stable, suspect tracking first (tags, website changes). Recent days can still grow because of conversion delay.'
  });
}

function checkConversionActions(ctx) {
  var ignore = CONVERSION_ACTIONS_IGNORE.map(function (n) { return String(n).toLowerCase(); });
  var actions = [];
  var rows = AdsApp.search(
    'SELECT conversion_action.name FROM conversion_action WHERE conversion_action.status = \'ENABLED\' ' +
    'AND conversion_action.include_in_conversions_metric = TRUE');
  while (rows.hasNext()) {
    var name = rows.next().conversionAction.name;
    if (ignore.indexOf(name.toLowerCase()) < 0) actions.push(name);
  }
  if (!actions.length) {
    addAlert(ctx, {
      severity: 'warning',
      title: 'No primary conversion action',
      message: 'The account has no enabled conversion action counted in the "Conversions" column.',
      hint: 'Smart bidding and every conversion check depend on this. Review Goals > Conversions.'
    });
    return;
  }
  var recent = period(addDays(ctx.today, -CONVERSION_ACTIONS_LOOKBACK_DAYS), ctx.yesterday);
  var recentConv = conversionsByAction(recent);
  var normalConv = conversionsByAction(ctx.baseline);
  var list = [];
  actions.forEach(function (name) {
    var expected = (normalConv[name] || 0) / ctx.baseline.days * recent.days;
    if (expected >= MIN_EXPECTED_CONVERSIONS && !(recentConv[name] > 0)) {
      list.push(name + ': 0 conversions (normally about ' + num(expected, 1) + ')');
    }
  });
  if (!list.length) return;
  addAlert(ctx, {
    severity: 'critical',
    title: list.length + ' primary conversion action(s) stopped recording',
    message: 'No conversions in ' + periodLabel(recent) + ' for actions that normally convert regularly.',
    list: list,
    hint: 'The tag of this action is probably broken. Check GTM, the thank-you page and the action\'s status in Goals > Conversions.'
  });
}

function conversionsByAction(p) {
  var out = {};
  var rows = AdsApp.search(
    'SELECT segments.conversion_action_name, metrics.conversions FROM campaign ' +
    'WHERE segments.date BETWEEN \'' + p.from + '\' AND \'' + p.to + '\' AND metrics.conversions > 0');
  while (rows.hasNext()) {
    var r = rows.next();
    var name = r.segments.conversionActionName;
    out[name] = (out[name] || 0) + Number(r.metrics.conversions || 0);
  }
  return out;
}

function checkConvValueMin(ctx) {
  if (ctx.w.value >= CONV_VALUE_MIN) return;
  addAlert(ctx, {
    severity: 'warning',
    title: 'Conversion value below minimum',
    message: 'Total conversion value in ' + periodLabel(ctx.window) + ' was below the set minimum.',
    facts: [['Conversion value', money(ctx.w.value)], ['Minimum', money(CONV_VALUE_MIN)], ['Spend in the same days', money(ctx.w.cost)]],
    hint: 'Check the purchase tag and whether the value is still passed in the data layer.'
  });
}

function checkConvValueVsAverage(ctx) {
  var expected = normalDaily(ctx.b, 'value') * ctx.w.n;
  if (expected <= 0) {
    Logger.log('Conversion value compared to normal: skipped, no conversion value in the normal period.');
    return;
  }
  if (!enoughVolume(ctx, 'Conversion value compared to normal')) return;
  var change = ctx.w.value / expected - 1;
  if (change >= -CONV_VALUE_MAX_DROP_PERCENT / 100) return;
  addAlert(ctx, {
    severity: 'warning',
    title: 'Conversion value dropped compared to normal',
    message: 'Conversion value is more than ' + CONV_VALUE_MAX_DROP_PERCENT + '% below ' + baselineLabel(ctx) + '.',
    facts: [['Conversion value', money(ctx.w.value)], ['Normal for ' + ctx.w.n + ' day(s)', money(expected)], ['Change', pct(change)]],
    hint: 'If the number of conversions is stable, the value is probably not passed correctly by the tag.'
  });
}

function checkRoas(ctx) {
  if (ctx.w.cost <= 0) return;
  if (ctx.b.value <= 0 && ctx.w.value <= 0) {
    Logger.log('ROAS: skipped, the account tracks no conversion value.');
    return;
  }
  if (!enoughVolume(ctx, 'ROAS')) return;
  var roas = ctx.w.value / ctx.w.cost;
  var ref, refLabel;
  if (ROAS_COMPARE_TO === 'TARGET') {
    ref = ROAS_TARGET_PERCENT / 100;
    refLabel = 'Target ROAS';
  } else {
    if (ctx.b.cost <= 0) return;
    ref = ctx.b.value / ctx.b.cost;
    refLabel = 'ROAS in the normal period';
  }
  if (ref <= 0) return;
  var change = roas / ref - 1;
  var limit = ROAS_MAX_DEVIATION_PERCENT / 100;
  var low = change < -limit && ROAS_DIRECTION !== 'ABOVE';
  var high = change > limit && ROAS_DIRECTION !== 'BELOW';
  if (!low && !high) return;
  addAlert(ctx, {
    severity: 'warning',
    title: low ? 'ROAS far below reference' : 'ROAS suspiciously high',
    message: 'ROAS in ' + periodLabel(ctx.window) + ' is more than ' + ROAS_MAX_DEVIATION_PERCENT + '% ' + (low ? 'below' : 'above') + ' the reference.',
    facts: [['ROAS', roasText(roas)], [refLabel, roasText(ref)], ['Difference', pct(change)], ['Spend', money(ctx.w.cost)], ['Conversion value', money(ctx.w.value)]],
    hint: low ? 'Check conversion tracking first, then search terms, bids and recent changes.'
      : 'Check for double-counted conversions or wrong values (e.g. currency or decimal errors in the tag).'
  });
}

function checkCpa(ctx) {
  if (ctx.w.conv <= 0 || ctx.b.conv <= 0) return;
  if (!enoughVolume(ctx, 'CPA compared to normal')) return;
  var cpa = ctx.w.cost / ctx.w.conv;
  var normal = ctx.b.cost / ctx.b.conv;
  var change = cpa / normal - 1;
  if (change <= CPA_MAX_INCREASE_PERCENT / 100) return;
  addAlert(ctx, {
    severity: 'warning',
    title: 'Cost per conversion (CPA) increased',
    message: 'CPA is more than ' + CPA_MAX_INCREASE_PERCENT + '% above ' + baselineLabel(ctx) + '.',
    facts: [['CPA', money(cpa)], ['Normal CPA', money(normal)], ['Change', pct(change)]],
    hint: 'Check conversion tracking, search terms, CPCs and landing page changes.'
  });
}

function checkDisapprovedAds(ctx) {
  var rows = AdsApp.search(
    'SELECT campaign.name, ad_group.name, ad_group_ad.ad.id FROM ad_group_ad ' +
    'WHERE ad_group_ad.policy_summary.approval_status = \'DISAPPROVED\' AND ad_group_ad.status = \'ENABLED\' ' +
    'AND ad_group.status = \'ENABLED\' AND campaign.status = \'ENABLED\'');
  var list = [];
  while (rows.hasNext()) {
    var r = rows.next();
    list.push(r.campaign.name + ' > ' + r.adGroup.name + ' (ad ID ' + r.adGroupAd.ad.id + ')');
  }
  if (!list.length) return;
  addAlert(ctx, {
    severity: 'warning',
    title: list.length + ' disapproved ad(s)',
    message: 'These ads are enabled but cannot run because of a policy disapproval.',
    list: list,
    hint: 'Open Ads, filter for "Disapproved", read the policy reason, then fix and resubmit.'
  });
}

function checkLostIsBudget(ctx) {
  var rows = AdsApp.search(
    'SELECT campaign.name, campaign_budget.amount_micros, metrics.cost_micros, metrics.search_impression_share, ' +
    'metrics.search_budget_lost_impression_share FROM campaign WHERE campaign.status = \'ENABLED\' ' +
    'AND campaign.advertising_channel_type = \'SEARCH\' AND metrics.cost_micros >= ' + Math.round(LOST_IS_MIN_SPEND * 1e6) +
    ' AND segments.date BETWEEN \'' + ctx.window.from + '\' AND \'' + ctx.window.to + '\'');
  var hits = [];
  while (rows.hasNext()) {
    var r = rows.next();
    if (r.metrics.searchBudgetLostImpressionShare == null) continue;
    var lost = Number(r.metrics.searchBudgetLostImpressionShare);
    if (lost * 100 <= LOST_IS_BUDGET_MAX_PERCENT) continue;
    hits.push({ lost: lost, text: r.campaign.name + ': ' + Math.round(lost * 100) + '% lost to budget' +
      (r.metrics.searchImpressionShare != null ? ', impression share ' + Math.round(Number(r.metrics.searchImpressionShare) * 100) + '%' : '') +
      ', daily budget ' + money(Number(r.campaignBudget.amountMicros || 0) / 1e6) + ', spend ' + money(Number(r.metrics.costMicros) / 1e6) });
  }
  if (!hits.length) return;
  hits.sort(function (a, b) { return b.lost - a.lost; });
  addAlert(ctx, {
    severity: 'warning',
    title: hits.length + ' Search campaign(s) losing impressions to budget',
    message: 'In ' + periodLabel(ctx.window) + ' these campaigns lost more than ' + LOST_IS_BUDGET_MAX_PERCENT +
      '% of their possible impressions because the daily budget ran out.',
    list: hits.map(function (h) { return h.text; }),
    hint: 'If the campaign is profitable, raise its budget or move budget from weaker campaigns. If not, lower bids or targets instead.'
  });
}

var BID_FIELDS = {
  CAMPAIGN: [
    ['campaign', 'targetRoas.targetRoas', 'Target ROAS', 'roas'],
    ['campaign', 'maximizeConversionValue.targetRoas', 'Target ROAS', 'roas'],
    ['campaign', 'targetCpa.targetCpaMicros', 'Target CPA', 'micros'],
    ['campaign', 'maximizeConversions.targetCpaMicros', 'Target CPA', 'micros']
  ],
  BIDDING_STRATEGY: [
    ['biddingStrategy', 'targetRoas.targetRoas', 'Portfolio Target ROAS', 'roas'],
    ['biddingStrategy', 'maximizeConversionValue.targetRoas', 'Portfolio Target ROAS', 'roas'],
    ['biddingStrategy', 'targetCpa.targetCpaMicros', 'Portfolio Target CPA', 'micros'],
    ['biddingStrategy', 'maximizeConversions.targetCpaMicros', 'Portfolio Target CPA', 'micros']
  ],
  AD_GROUP: [
    ['adGroup', 'cpcBidMicros', 'Ad group max. CPC', 'micros'],
    ['adGroup', 'targetCpaMicros', 'Ad group Target CPA', 'micros'],
    ['adGroup', 'targetRoas', 'Ad group Target ROAS', 'roas']
  ],
  AD_GROUP_CRITERION: [
    ['adGroupCriterion', 'cpcBidMicros', 'Keyword max. CPC', 'micros']
  ]
};

var CLIENT_TYPES = {
  GOOGLE_ADS_WEB_CLIENT: 'Google Ads web interface',
  GOOGLE_ADS_AUTOMATED_RULE: 'automated rule',
  GOOGLE_ADS_SCRIPTS: 'Google Ads script',
  GOOGLE_ADS_BULK_UPLOAD: 'bulk upload',
  GOOGLE_ADS_API: 'API / external tool',
  GOOGLE_ADS_EDITOR: 'Google Ads Editor',
  GOOGLE_ADS_MOBILE_APP: 'Google Ads mobile app',
  GOOGLE_ADS_RECOMMENDATIONS: 'Google recommendation',
  GOOGLE_ADS_RECOMMENDATIONS_SUBSCRIPTION: 'auto-applied Google recommendation',
  INTERNAL_TOOL: 'Google internal tool'
};

function checkBudgetChanges(ctx) {
  var ch = loadChanges(ctx);
  var list = [];
  ch.events.forEach(function (e) {
    if (e.type !== 'CAMPAIGN_BUDGET') return;
    var before = getPath(e.old, 'campaignBudget.amountMicros');
    var after = getPath(e.new, 'campaignBudget.amountMicros');
    if (before == null || after == null || Number(before) <= 0) return;
    var change = Number(after) / Number(before) - 1;
    if (Math.abs(change) * 100 < BUDGET_CHANGE_PERCENT) return;
    var campaigns = ch.budgetCampaigns[e.resource] || [];
    list.push((campaigns.length ? campaigns.join(', ') : 'Shared budget') + ': ' +
      money(Number(before) / 1e6) + ' -> ' + money(Number(after) / 1e6) + ' (' + pct(change) + ')' + changeSuffix(e));
  });
  if (!list.length) return;
  addAlert(ctx, {
    severity: 'warning',
    title: list.length + ' big budget change(s)',
    message: 'Campaign budgets changed by at least ' + BUDGET_CHANGE_PERCENT + '% in ' + periodLabel(ch.period) + '.',
    list: list,
    hint: 'Make sure each change was intended and agreed with the client, and that the monthly budget still fits.'
  });
}

function checkBidChanges(ctx) {
  var ch = loadChanges(ctx);
  var list = [];
  ch.events.forEach(function (e) {
    var where = changeTarget(ch, e);
    if (e.type === 'CAMPAIGN') {
      var oldType = getPath(e.old, 'campaign.biddingStrategyType');
      var newType = getPath(e.new, 'campaign.biddingStrategyType');
      if (oldType && newType && oldType !== newType) {
        list.push(where + ': bid strategy ' + readable(oldType) + ' -> ' + readable(newType) + changeSuffix(e));
      }
    }
    (BID_FIELDS[e.type] || []).forEach(function (f) {
      var before = getPath(e.old, f[0] + '.' + f[1]);
      var after = getPath(e.new, f[0] + '.' + f[1]);
      if (before == null || after == null || Number(before) <= 0) return;
      var change = Number(after) / Number(before) - 1;
      if (Math.abs(change) * 100 < BID_CHANGE_PERCENT) return;
      var show = f[3] === 'roas' ? roasText : function (v) { return money(v / 1e6); };
      list.push(where + ': ' + f[2] + ' ' + show(Number(before)) + ' -> ' + show(Number(after)) + ' (' + pct(change) + ')' + changeSuffix(e));
    });
  });
  if (!list.length) return;
  addAlert(ctx, {
    severity: 'warning',
    title: list.length + ' big bid change(s)',
    message: 'Bids or bidding targets changed by at least ' + BID_CHANGE_PERCENT + '%, or a bid strategy was switched, in ' + periodLabel(ch.period) + '.',
    list: list,
    hint: 'Large target changes restart Smart Bidding learning and can move spend and ROAS sharply. Confirm they were intended.'
  });
}

function loadChanges(ctx) {
  if (ctx.changes) return ctx.changes;
  var p = period(addDays(ctx.today, -CHANGE_HISTORY_DAYS), ctx.yesterday);
  var ignore = CHANGE_HISTORY_IGNORE_USERS.map(function (u) { return String(u).toLowerCase().trim(); });
  var ch = { period: p, events: [], campaigns: {}, budgetCampaigns: {}, adGroups: {}, strategies: {} };

  var rows = AdsApp.search(
    'SELECT change_event.change_date_time, change_event.change_resource_type, change_event.change_resource_name, ' +
    'change_event.old_resource, change_event.new_resource, change_event.user_email, change_event.client_type, ' +
    'change_event.campaign, change_event.ad_group FROM change_event ' +
    'WHERE change_event.change_date_time >= \'' + p.from + ' 00:00:00\' AND change_event.change_date_time <= \'' + p.to + ' 23:59:59\' ' +
    'AND change_event.change_resource_type IN (\'CAMPAIGN_BUDGET\', \'CAMPAIGN\', \'AD_GROUP\', \'AD_GROUP_CRITERION\', \'BIDDING_STRATEGY\') ' +
    'AND change_event.resource_change_operation = \'UPDATE\' ORDER BY change_event.change_date_time ASC LIMIT 10000');
  while (rows.hasNext()) {
    var c = rows.next().changeEvent;
    if (c.userEmail && ignore.indexOf(c.userEmail.toLowerCase()) >= 0) continue;
    ch.events.push({
      time: c.changeDateTime, type: c.changeResourceType, resource: c.changeResourceName,
      old: c.oldResource || {}, new: c.newResource || {},
      user: c.userEmail, client: c.clientType, campaign: c.campaign, adGroup: c.adGroup
    });
  }
  if (!ch.events.length) return (ctx.changes = ch);

  rows = AdsApp.search('SELECT campaign.resource_name, campaign.name, campaign.campaign_budget FROM campaign');
  while (rows.hasNext()) {
    var cp = rows.next().campaign;
    ch.campaigns[cp.resourceName] = cp.name;
    if (cp.campaignBudget) (ch.budgetCampaigns[cp.campaignBudget] = ch.budgetCampaigns[cp.campaignBudget] || []).push(cp.name);
  }
  if (ch.events.some(function (e) { return e.type === 'BIDDING_STRATEGY'; })) {
    rows = AdsApp.search('SELECT bidding_strategy.resource_name, bidding_strategy.name FROM bidding_strategy');
    while (rows.hasNext()) {
      var bs = rows.next().biddingStrategy;
      ch.strategies[bs.resourceName] = bs.name;
    }
  }
  var adGroupIds = {};
  ch.events.forEach(function (e) { if (e.adGroup) adGroupIds[e.adGroup] = true; });
  var names = Object.keys(adGroupIds);
  for (var i = 0; i < names.length; i += 500) {
    rows = AdsApp.search('SELECT ad_group.resource_name, ad_group.name FROM ad_group WHERE ad_group.resource_name IN (\'' +
      names.slice(i, i + 500).join('\', \'') + '\')');
    while (rows.hasNext()) {
      var ag = rows.next().adGroup;
      ch.adGroups[ag.resourceName] = ag.name;
    }
  }
  return (ctx.changes = ch);
}

function changeTarget(ch, e) {
  if (e.type === 'BIDDING_STRATEGY') return 'Portfolio strategy "' + (ch.strategies[e.resource] || e.resource) + '"';
  var parts = [ch.campaigns[e.campaign] || 'Unknown campaign'];
  if (e.adGroup) parts.push(ch.adGroups[e.adGroup] || 'Unknown ad group');
  return parts.join(' > ');
}

function changeSuffix(e) {
  var via = CLIENT_TYPES[e.client] || readable(e.client || 'unknown');
  return ' | ' + String(e.time).substring(0, 16) + ', by ' + (e.user || via) + (e.user ? ' via ' + via : '');
}

function getPath(obj, path) {
  return path.split('.').reduce(function (o, k) { return o == null ? undefined : o[k]; }, obj);
}

function readable(enumValue) {
  return String(enumValue).toLowerCase().replace(/_/g, ' ');
}

// ---------- Email ----------

var SEVERITY = {
  critical: { label: 'CRITICAL', color: '#c5221f' },
  warning: { label: 'WARNING', color: '#e37400' },
  info: { label: 'INFO', color: '#1a73e8' }
};

function addAlert(ctx, alert) {
  ctx.alerts.push(alert);
  Logger.log('[' + SEVERITY[alert.severity].label + '] ' + alert.title + ': ' + alert.message);
}

function recordError(ctx, check, e) {
  var message = e && e.message ? e.message : String(e);
  ctx.errors.push({ check: check, message: message, stack: e && e.stack ? String(e.stack) : '' });
  Logger.log('[ERROR] ' + check + ': ' + message);
}

function sendReport(ctx) {
  var account = ctx.account || safeAccount();
  var order = { critical: 0, warning: 1, info: 2 };
  ctx.alerts.sort(function (a, b) { return order[a.severity] - order[b.severity]; });

  if (!ctx.alerts.length && !ctx.errors.length) {
    Logger.log('All guardrails passed.');
    if (!SEND_ALL_CLEAR_EMAIL) return;
  }

  var parts = [];
  if (ctx.errors.length) parts.push('SCRIPT ERROR');
  var critical = ctx.alerts.filter(function (a) { return a.severity === 'critical'; }).length;
  if (critical) parts.push(critical + ' critical');
  if (ctx.alerts.length - critical) parts.push((ctx.alerts.length - critical) + ' warning');
  if (!parts.length) parts.push('All clear');
  var preview = AdsApp.getExecutionInfo().isPreview() ? '[PREVIEW] ' : '';
  var subject = preview + '[GAds Guardrail] ' + parts.join(', ') + ' | ' + account.name + ' (' + account.id + ')';

  var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#202124;max-width:720px">' +
    '<h2 style="margin:0 0 4px">Google Ads guardrail report</h2>' +
    '<p style="margin:0 0 16px;color:#5f6368">' + esc(account.name) + ' (' + esc(account.id) + ')' +
    (ctx.window ? '<br>Checked: ' + esc(periodLabel(ctx.window)) + '<br>Compared with: ' + esc(baselineLabel(ctx)) : '') + '</p>';
  var text = 'Google Ads guardrail report\n' + account.name + ' (' + account.id + ')\n' +
    (ctx.window ? 'Checked: ' + periodLabel(ctx.window) + '\nCompared with: ' + baselineLabel(ctx) + '\n' : '') + '\n';

  ctx.errors.forEach(function (e) {
    html += card('#5f6368', 'SCRIPT ERROR', 'Check failed: ' + e.check,
      'This check could not run, so its guardrail is NOT active until the problem is fixed.',
      [['Error', e.message]], [], 'Open the script in Google Ads and look at the Logs. If the error repeats, the script or a setting needs fixing.');
    text += '[SCRIPT ERROR] Check failed: ' + e.check + '\n' + e.message + '\n' + e.stack + '\n\n';
  });
  ctx.alerts.forEach(function (a) {
    var s = SEVERITY[a.severity];
    html += card(s.color, s.label, a.title, a.message, a.facts || [], a.list || [], a.hint);
    var lines = ['[' + s.label + '] ' + a.title, a.message]
      .concat((a.facts || []).map(function (f) { return '  ' + f[0] + ': ' + f[1]; }))
      .concat((a.list || []).slice(0, MAX_LIST_ITEMS).map(function (l) { return '  - ' + l; }));
    if (a.hint) lines.push('What to do: ' + a.hint);
    text += lines.join('\n') + '\n\n';
  });
  if (!ctx.alerts.length && !ctx.errors.length) {
    html += '<p>Every enabled guardrail passed.</p>';
    text += 'Every enabled guardrail passed.\n';
  }
  html += '<p style="color:#80868b;font-size:12px;margin-top:24px">Sent by the Google Ads guardrail script. Settings can be changed in the SETUP section of the script.</p></div>';

  MailApp.sendEmail({ to: EMAIL_RECIPIENTS, subject: subject, body: text, htmlBody: html });
  Logger.log('Email sent to ' + EMAIL_RECIPIENTS + ': ' + subject);
}

function card(color, label, title, message, facts, list, hint) {
  var h = '<div style="border-left:4px solid ' + color + ';background:#f8f9fa;padding:12px 16px;margin:0 0 12px">' +
    '<div style="font-size:11px;font-weight:bold;color:' + color + '">' + label + '</div>' +
    '<div style="font-size:16px;font-weight:bold;margin:2px 0 6px">' + esc(title) + '</div>' +
    '<div>' + esc(message) + '</div>';
  if (facts.length) {
    h += '<table style="border-collapse:collapse;margin-top:8px">' + facts.map(function (f) {
      return '<tr><td style="padding:2px 16px 2px 0;color:#5f6368">' + esc(f[0]) + '</td><td style="padding:2px 0;font-weight:bold">' + esc(f[1]) + '</td></tr>';
    }).join('') + '</table>';
  }
  if (list.length) {
    h += '<ul style="margin:8px 0 0;padding-left:20px">' + list.slice(0, MAX_LIST_ITEMS).map(function (l) {
      return '<li>' + esc(l) + '</li>';
    }).join('') + (list.length > MAX_LIST_ITEMS ? '<li>... and ' + (list.length - MAX_LIST_ITEMS) + ' more</li>' : '') + '</ul>';
  }
  if (hint) h += '<div style="margin-top:8px;color:#3c4043"><b>What to do:</b> ' + esc(hint) + '</div>';
  return h + '</div>';
}

function safeAccount() {
  try {
    var a = AdsApp.currentAccount();
    return { name: ACCOUNT_LABEL || a.getName(), id: a.getCustomerId() };
  } catch (e) {
    return { name: ACCOUNT_LABEL || 'Unknown account', id: '?' };
  }
}

// ---------- Helpers ----------

function period(start, end) {
  return { start: start, end: end, from: ymd(start), to: ymd(end), days: Math.round((end - start) / DAY_MS) + 1 };
}

function periodLabel(p) {
  return p.days === 1 ? p.from : p.from + ' to ' + p.to + ' (' + p.days + ' days)';
}

function baselineLabel(ctx) {
  var how = BASELINE_AVERAGE === 'MEDIAN' ? 'median day of ' : 'daily average of ';
  return how + (BASELINE_PERIOD === 'LAST_MONTH' ? 'last month' : 'the previous ' + BASELINE_DAYS + ' days') +
    ' (' + ctx.baseline.from + ' to ' + ctx.baseline.to + ')';
}

function parseYmd(s) {
  var p = s.split('-');
  return new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12));
}

function addDays(d, n) {
  return new Date(d.getTime() + n * DAY_MS);
}

function ymd(d) {
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function num(v, decimals) {
  var parts = Math.abs(v).toFixed(decimals).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return (v < 0 ? '-' : '') + parts.join('.');
}

function money(v) {
  return num(v, CURRENCY === 'HUF' ? 0 : 2) + ' ' + CURRENCY;
}

function pct(ratio) {
  return (ratio >= 0 ? '+' : '') + Math.round(ratio * 100) + '%';
}

function roasText(ratio) {
  return Math.round(ratio * 100) + '%';
}

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
