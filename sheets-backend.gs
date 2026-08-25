/**
 * FinLit Index · Pune 2026 — Google Sheets collector
 *
 * Spreadsheet:
 * https://docs.google.com/spreadsheets/d/1cntlU1VebqoA2AZE9Vqa1VQ6st-jW8xYNZwVD-Yp044/edit
 *
 * INSTALL (once)
 * 1. Open that spreadsheet
 * 2. Extensions → Apps Script
 * 3. Replace ALL code with this file and Save
 * 4. Select fixShiftedRowsNow → Run
 *    Review permissions if a popup appears, then look at the spreadsheet
 *    (not the script tab) for a brief FinLit toast. Do not run setup first.
 * 5. Deploy → Manage deployments → pencil on the Web app
 *    Version: New version → Deploy
 *
 * Rebuild dashboard is optional and slower (charts). Use FinLit → Rebuild
 * dashboard from the spreadsheet menu when you want charts refreshed.
 *
 * Scores come from the quiz itself. Running setup does not rescore old rows.
 * Every submission must include a real email. That address is stored in
 * column C on Responses and is mailed the score, missed questions, and
 * explanations (newspaper-style HTML). First Gmail run asks for permission.
 * for permission. Deploy → New version or the live form keeps using old code.
 * If scores were wiped to 0, restore a Google Sheets version from before that
 * happened: File → Version history → See version history.
 *
 * Q22 key is A (₹50 lakh). Q25 key is B. Profession and correct-count are not stored.
 */

var KEYS = [
  "B","B","C","B","C",
  "B","C","D","B","C","B","C",
  "C","B","C","D","C","C","C","A",
  "C","A","B","C","B","B","B"
];

var MARKS = [
  0.5,0.5,0.5,0.5,0.5,
  1,1,1,1,1,1,1,
  1.5,1.5,1.5,1.5,1.5,1.5,1.5,1.5,
  0.5,0.5,0.5,0.5,0.5,0.5,0.5
];

// BEGIN FEEDBACK_BANK
var FEEDBACK_BANK = [
  {
    n: 1, part: "A", marks: 0.5, key: "B",
    text: "Meera has ₹800 in pocket money this month. She has four possible expenses:",
    bullets: ["Exam stationery she needs this week (₹150)","College application fee due Friday, after which she loses her spot (₹250)","A new phone case she has been wanting (₹300)","A birthday gift for her best friend next week (₹200)"],
    prompt: "She cannot afford all four. What should she pay for first?",
    opts: {A:"The phone case, since it is the one she wants most", B:"The application fee, since missing that deadline cannot be reversed", C:"The birthday gift, so her friend is not upset", D:"Split the ₹800 evenly across everything"},
    why: "When money is short, pay first for the cost that cannot be undone. Missing Friday’s application deadline loses the college spot for good. Stationery can be bought a little later, a gift can wait a few days, and a phone case is a want. Splitting the money evenly would still leave the fee unpaid. Needs and irreversible deadlines come before wants."
  },
  {
    n: 2, part: "A", marks: 0.5, key: "B",
    text: "Rohan gets a text: \"Your electricity bill payment failed. Pay ₹49 immediately via this link to keep your power on tonight: bit.ly/xyz123.\"",
    bullets: [],
    prompt: "What is the biggest sign that this message is a scam?",
    opts: {A:"The amount asked for is very small", B:"It uses an unknown short link and pushes him to act tonight", C:"Electricity messages usually arrive by email, not SMS", D:"The bill amount seems too low"},
    why: "Scam messages try to skip thinking. An unknown shortened link hides the real website, and “pay tonight or the power goes off” is pressure. A small amount is not proof of a scam — thieves often ask for a little so you do not hesitate. Utilities can message by SMS, and a low-looking bill is not the main clue. Never tap a surprise payment link. Open the official app or website yourself."
  },
  {
    n: 3, part: "A", marks: 0.5, key: "C",
    text: "Ayesha's grandmother gives her ₹5,000 in cash for her birthday. She keeps it in a locked drawer at home.",
    bullets: [],
    prompt: "What is the main advantage of putting this money into a bank savings account instead?",
    opts: {A:"A passbook entry proves she has the money", B:"Money in a bank is harder to spend on impulse", C:"The money is safer and earns some interest over time", D:"Banks double the deposit after a few years"},
    why: "Cash in a drawer can be stolen, lost, or destroyed. A savings account is protected by the bank and, in India, deposit insurance up to the covered limit. It also pays interest, so the money grows a little while it sits. A passbook is only a record. Banks do not double ordinary savings after a few years. Impulse control can help, but safety plus interest is the main reason to bank the cash."
  },
  {
    n: 4, part: "A", marks: 0.5, key: "B",
    text: "Karan puts ₹2,000 into an account that pays 5% simple interest per year.",
    bullets: [],
    prompt: "How much interest will he have earned after 2 years?",
    opts: {A:"₹100", B:"₹200", C:"₹205", D:"₹400"},
    why: "Simple interest is charged only on the original amount, not on interest already earned. Interest each year is 5% of ₹2,000, which is ₹100. Over two years that is ₹100 × 2 = ₹200. ₹100 would be one year only. ₹205 would be as if a little compounding had been added. ₹400 would be 10% a year, or 5% of the wrong principal."
  },
  {
    n: 5, part: "A", marks: 0.5, key: "C",
    text: "An emergency fund is money kept aside for situations you did not plan for.",
    bullets: [],
    prompt: "Which of these is the main problem an emergency fund is meant to solve?",
    opts: {A:"The risk that your savings lose value to rising prices", B:"The urge to spend on things you don't really need", C:"A sudden expense like a hospital bill or urgent repair, when you have nothing else to fall back on", D:"The chance of missing out on stock market gains"},
    why: "An emergency fund is a cash cushion for shocks: illness, a broken phone, a parent losing work. Without it, people often borrow at high cost. Inflation is a separate problem (investing). Impulse spending is a budgeting habit. Missing stock-market gains is not what this money is for — emergency money should stay easy to reach, not locked in shares."
  },
  {
    n: 6, part: "B", marks: 1, key: "B",
    text: "In 2015, a filter coffee in Pune cost ₹15. Today, the same coffee costs ₹30. Ravi has kept ₹1,500 in cash at home, untouched, since 2015.",
    bullets: [],
    prompt: "What has happened to what his ₹1,500 can actually buy?",
    opts: {A:"Nothing changed. It is still ₹1,500", B:"It now buys roughly half of what it did in 2015", C:"It has grown, since prices going up means money is worth more", D:"It depends only on where he shops"},
    why: "The note still says ₹1,500, but prices doubled, so each rupee buys about half as much. That is inflation: the same cash, less purchasing power. Prices going up means money is worth less, not more. Shopping around can change a bill a little, but it does not restore 2015 prices."
  },
  {
    n: 7, part: "B", marks: 1, key: "C",
    text: "Sana puts ₹10,000 into a bank for 3 years. Bank A pays 8% simple interest each year. Bank B pays 8% each year, but each year's interest gets added to the amount and starts earning interest itself (this is called compound interest).",
    bullets: [],
    prompt: "Which bank leaves Sana with more money at the end?",
    opts: {A:"Both give the same result, since the rate is 8% either way", B:"Bank A, because the calculation is simpler each year", C:"Bank B, because interest keeps building on interest", D:"Bank A, because compounding takes something away each year"},
    why: "The headline rate is 8% at both banks. The difference is whether last year’s interest itself earns interest. Bank B compounds, so the balance grows faster. Simple interest (Bank A) pays 8% only on the original ₹10,000 every year. Compounding does not take money away; it adds extra growth."
  },
  {
    n: 8, part: "B", marks: 1, key: "D",
    text: "Aditya has ₹50,000 that he is sure he will not need for the next 3 years. He wants the best guaranteed return with no risk of loss.",
    bullets: [],
    prompt: "Which option best suits what he wants?",
    opts: {A:"A regular savings account, so he can withdraw any time", B:"A recurring deposit, adding a small amount each month", C:"Buying shares of a big company, since 3 years is enough time", D:"A fixed deposit locked in for the full 3 years"},
    why: "He already has the lump sum, will not need it for three years, and wants a guaranteed return with no chance of loss. A fixed deposit locked for those three years usually pays more than a savings account and does not fall in value the way shares can. A recurring deposit is for people adding money each month, not for parking ₹50,000 already in hand. Shares can lose money even in three years."
  },
  {
    n: 9, part: "B", marks: 1, key: "B",
    text: "Priyanka gets a ₹15,000 scholarship and spends all of it on branded sneakers. Her friend Neha gets the same amount and puts it in a bank account that pays 7% interest per year for 2 years.",
    bullets: [],
    prompt: "Beyond the ₹15,000 itself, what has Priyanka's decision effectively cost her over those two years?",
    opts: {A:"Nothing extra. The sneakers still exist", B:"The interest her money could have earned in the bank, which Neha will have and she will not", C:"The GST portion of the sneaker price", D:"Only the resale value the sneakers have lost"},
    why: "Opportunity cost is what you give up by choosing one thing. Priyanka gave up the interest Neha earns — about 7% a year on ₹15,000 for two years. The sneakers still exist, but that does not cancel the missed growth. GST and resale value are separate from the interest she will never receive."
  },
  {
    n: 10, part: "B", marks: 1, key: "C",
    text: "Nikhil uses UPI every day for auto rides, groceries, and recharges. He feels very confident about handling money. He has never checked what interest his savings account pays, and he does not know what his expenses add up to in a month.",
    bullets: [],
    prompt: "What does his situation best show?",
    opts: {A:"Using UPI daily is enough to build financial understanding", B:"His confidence proves he is managing his money well", C:"Being comfortable with paying is not the same as understanding money", D:"Tracking expenses is not needed once payments are digital"},
    why: "Paying with UPI is a tool, like knowing how to swipe a card. Financial literacy is knowing what you earn, what you spend, what interest you earn or pay, and whether you are getting ahead. Nikhil is fluent at sending money and still blind to his own numbers. Confidence without those facts is not the same as skill."
  },
  {
    n: 11, part: "B", marks: 1, key: "B",
    text: "Yash puts his entire ₹3,00,000 in savings into shares of one small tech company, because a relative told him the stock doubled last year.",
    bullets: [],
    prompt: "What is the most serious risk in his decision?",
    opts: {A:"Small tech companies always carry more risk than the wider market", B:"One bad year at this one company could wipe out most of his savings", C:"A stock that doubles one year usually falls the next", D:"A relative's tip is less reliable than a broker's"},
    why: "The grave problem is concentration: all his savings ride on one company. If that firm has a bad year, most of the ₹3,00,000 can disappear. Small tech stocks are often riskier, but even a large company can fail. Last year’s doubling does not mean a fall next year, and a broker’s tip can be as biased as a relative’s. Spreading money across many investments is how you avoid a single story wiping you out."
  },
  {
    n: 12, part: "B", marks: 1, key: "C",
    text: "Fatima is 29, self employed, and has ₹60,000 in her savings account. She has no health insurance from a company or from the government. Suppose she falls seriously ill and needs two weeks in a private hospital.",
    bullets: [],
    prompt: "What is health insurance actually meant to protect her from?",
    opts: {A:"It guarantees she never falls ill", B:"It reduces the price of the treatment for everyone in the country", C:"It covers the hospital bill, which could easily be larger than her entire savings, so one illness does not wipe her out", D:"It replaces her lost income while she recovers"},
    why: "Health insurance does not stop illness and does not set hospital prices for the whole country. A private stay of two weeks can cost far more than ₹60,000. The policy is meant to pay that bill so one event does not empty her savings. Lost income is what disability or income-protection cover is for, not a standard health policy."
  },
  {
    n: 13, part: "C", marks: 1.5, key: "C",
    text: "Deepak wants to invest ₹1,20,000 in a share based mutual fund this year. He is worried about putting it all in right before a possible market fall.",
    bullets: [],
    prompt: "Which approach best reduces the risk of poor timing?",
    opts: {A:"Invest the full amount today and stop worrying about timing", B:"Keep the money in cash until the next big market fall", C:"Spread it out as ₹10,000 every month for a year", D:"Split it evenly between gold and the fund"},
    why: "Putting ₹10,000 in each month (a SIP) buys more units when prices are low and fewer when they are high, so one unlucky day does not decide the whole ₹1,20,000. Investing everything today maximises timing risk. Waiting for “the next crash” often means sitting in cash for years. Splitting with gold changes the mix; it does not solve the timing of the equity purchase."
  },
  {
    n: 14, part: "C", marks: 1.5, key: "B",
    text: "A shop offers a phone on a 'zero interest' EMI of ₹4,200 per month for 12 months. The same phone costs ₹45,000 if you pay in cash.",
    bullets: [],
    prompt: "What is actually true about this offer?",
    opts: {A:"It is genuinely free financing, since no interest is charged", B:"The 12 EMIs add up to ₹50,400, which is ₹5,400 more than the cash price", C:"Paying by EMI is always cheaper, since your cash stays in the bank", D:"Both options cost you the same amount in the end"},
    why: "Always multiply. ₹4,200 × 12 = ₹50,400. Cash is ₹45,000. The EMI path costs ₹5,400 extra, whatever the poster says about “zero interest.” That extra is the cost of spreading payments — often built into a higher EMI price. EMI is not always cheaper, and the two options do not cost the same."
  },
  {
    n: 15, part: "C", marks: 1.5, key: "C",
    text: "Mr. Iyer is 55 and plans to retire in 5 years. He has 90% of his retirement savings in one type of fund that invests in small companies. These funds gave the highest returns of any product over the last ten years, but their value can swing sharply from year to year.",
    bullets: [],
    prompt: "What is the biggest concern with his portfolio right now?",
    opts: {A:"Nothing. His funds have the best track record", B:"He does not hold enough of these funds for a 5 year horizon", C:"If the market falls sharply near his retirement, he will have very little time to recover before he needs the money", D:"This type of fund is not allowed for investors above 55"},
    why: "Past returns do not protect the next five years. Small-company funds can drop hard. At 55, he may need this money soon. A crash just before retirement can force him to sell at a low price with no years left to wait for a rebound. He already holds 90% in this risky sleeve, which is too much, not too little. Age 55 does not ban these funds; it makes a crash more dangerous."
  },
  {
    n: 16, part: "C", marks: 1.5, key: "D",
    text: "Meenal puts ₹1,00,000 into a fixed deposit that grows at 7% per year, for 5 years. Over the same 5 years, prices in general go up by 6% per year (this is inflation).",
    bullets: [],
    prompt: "What is really happening to her money's buying power?",
    opts: {A:"It grows well, since 7 is comfortably above 6", B:"She loses buying power every year she stays in it", C:"Since the rate is fixed, inflation has no effect on the outcome", D:"She gains only about 1% per year in real buying power"},
    why: "The bank credits 7%, but prices rise 6%. The real gain is about 7% − 6% = 1% a year. She is not going backwards, and 7 is not “comfortably” large once inflation is counted. A fixed rate does not freeze prices in the shops. Always subtract inflation to see what the money can actually buy."
  },
  {
    n: 17, part: "C", marks: 1.5, key: "C",
    text: "You want to buy a second hand laptop for ₹30,000 to try video editing on the side, which might earn you some money later. You save ₹2,500 a month from your allowance. You do not need the laptop urgently.",
    bullets: [],
    prompt: "What is the most sensible way to think about this purchase?",
    opts: {A:"Take a small loan now, so the laptop can start earning for you sooner", B:"Put it on a credit card and pay it off in bits later", C:"Save for a year and buy it in cash, so you can find out whether the hobby actually earns anything without any debt on top", D:"Ask a parent to buy it as a gift and save your allowance for something else"},
    why: "The laptop is optional and the income is only a “might.” Borrowing (loan or credit card) adds interest and a repayment you must meet even if editing earns nothing. ₹2,500 × 12 = ₹30,000, so a year of saving buys it in cash and tests the hobby with no debt. Asking a parent to pay does not teach you whether the work covers the cost."
  },
  {
    n: 18, part: "C", marks: 1.5, key: "C",
    text: "Ravi has finished Class 12 and is choosing between two options. Path 1 is a 4 year engineering degree costing ₹8 lakh in total, with typical starting salaries of ₹6 lakh a year. Path 2 is a 1 year technical course costing ₹1.5 lakh, with typical starting salaries of ₹4 lakh a year.",
    bullets: [],
    prompt: "Purely as a financial decision, what is the most complete way to compare the two?",
    opts: {A:"Whichever costs less overall is the better option", B:"Whichever pays more once you graduate is the better option", C:"The total cost, plus how soon you start earning, plus what you earn over five to ten years, all weighed together", D:"Whichever qualification sounds more impressive when you apply for jobs"},
    why: "A cheaper course can still lose if you earn less for years. A higher starting salary can still lose if you spent four extra years not earning and paid ₹8 lakh. The full comparison is cost, when pay starts, and what you earn over a stretch of years — not prestige, and not cost or starting pay alone."
  },
  {
    n: 19, part: "C", marks: 1.5, key: "C",
    text: "Aarav is 25. Kamlesh is 58. Each has ₹5,00,000 to invest. Both put 80% into shares (higher risk, higher possible return) and 20% into bonds (lower risk, steady return).",
    bullets: [],
    prompt: "Why is this same split much riskier for Kamlesh than for Aarav?",
    opts: {A:"Older people always earn less than younger people", B:"Share prices fall faster the older you get", C:"Aarav has 30 or more years to recover if shares crash. Kamlesh may need this money in the next few years, so a crash near then would hurt him badly", D:"Bonds are legally required to make up more of the portfolio after age 50"},
    why: "Shares do not fall faster because you are older. The clock is different. Aarav can wait decades after a crash. Kamlesh may retire soon and need the money, so a fall would force sales at a bad time. There is no law that bonds must rise after age 50 — it is a time-horizon choice."
  },
  {
    n: 20, part: "C", marks: 1.5, key: "A",
    text: "The Kulkarnis save ₹10,000 every month in a savings account that grows at 3% a year. Prices are going up by 6% a year. They want to save ₹15,00,000 in 10 years for their child's college, using today's college prices as their target.",
    bullets: [],
    prompt: "What is likely to go wrong with their plan?",
    opts: {A:"The money is growing slower than prices are rising, and college itself will cost more in 10 years than it does today", B:"The plan is fine. ₹10,000 a month is just too little", C:"Given ten years, the growth will overtake rising prices on its own", D:"Everything works out if they simply delay the goal by two years"},
    why: "Two problems stack. The account grows at 3% while prices rise at 6%, so the savings lose buying power every year. College fees will also be higher in ten years than today’s ₹15,00,000 target. Ten years does not magically make 3% beat 6%. Delaying two years makes the fee inflation worse, not better. They need a higher-return mix and a target that includes future fees."
  },
  {
    n: 21, part: "D", marks: 0.5, key: "C",
    text: "What is the minimum amount someone needs to invest in an Alternative Investment Fund (AIF) in India?",
    bullets: [],
    prompt: "",
    opts: {A:"₹10 lakh", B:"₹50 lakh", C:"₹1 crore", D:"₹5 crore"},
    why: "SEBI sets a high floor so AIFs stay a product for large investors. The usual minimum commitment is ₹1 crore. ₹10 lakh and ₹50 lakh are below that bar (₹50 lakh is the PMS floor). ₹5 crore is higher than the standard minimum."
  },
  {
    n: 22, part: "D", marks: 0.5, key: "A",
    text: "What is the minimum amount someone needs to invest in a Portfolio Management Service (PMS) in India?",
    bullets: [],
    prompt: "",
    opts: {A:"₹50 lakh", B:"₹25 lakh", C:"₹1 crore", D:"₹10 lakh"},
    why: "SEBI’s current minimum for a PMS account is ₹50 lakh. That is lower than the AIF floor of ₹1 crore and higher than ₹10 lakh or ₹25 lakh. A PMS is a professionally managed portfolio for people who can put in that much capital."
  },
  {
    n: 23, part: "D", marks: 0.5, key: "B",
    text: "What is the most common minimum monthly amount to start an SIP in a mutual fund?",
    bullets: [],
    prompt: "",
    opts: {A:"₹100", B:"₹500", C:"₹1,000", D:"₹5,000"},
    why: "A systematic investment plan (SIP) lets you put in a fixed amount every month. Many funds now allow ₹100, but the amount most often used as the standard starting SIP is ₹500. ₹1,000 and ₹5,000 are common later, not the usual minimum to begin."
  },
  {
    n: 24, part: "D", marks: 0.5, key: "C",
    text: "Which of these has grown the most in India over the last 30 years?",
    bullets: [],
    prompt: "",
    opts: {A:"Gold", B:"Real estate", C:"Shares of Indian companies", D:"Bank fixed deposits"},
    why: "Over very long stretches (around 30 years), a diversified basket of Indian company shares has compounded faster than gold, typical city property, or bank deposits. Shares bounce around year to year, which is why this is a long-horizon fact, not a promise for the next two years. Fixed deposits barely beat inflation after tax. Gold and property have done well in some periods, but listed equities have led across three decades."
  },
  {
    n: 25, part: "D", marks: 0.5, key: "B",
    text: "Which of these has grown the most in India over the last 10 years?",
    bullets: [],
    prompt: "",
    opts: {A:"Bank fixed deposits", B:"Shares of Indian companies", C:"Real estate", D:"Gold"},
    why: "On a ten-year lookback, Indian equities (broad share indices) have generally outpaced gold, typical residential property, and bank fixed deposits. Ten years is long enough for shares to show that lead, though any single decade can differ. FDs stay near the interest rate. Gold and property vary by city and cycle."
  },
  {
    n: 26, part: "D", marks: 0.5, key: "B",
    text: "Which of these has grown the most in India over the last 2 years?",
    bullets: [],
    prompt: "",
    opts: {A:"Shares of Indian companies", B:"Gold", C:"Real estate", D:"Bank fixed deposits"},
    why: "Short windows flip the ranking. In the most recent two-year stretch used for this paper, gold led, helped by global uncertainty and a strong bullion rally, while Indian shares were more mixed and deposits stayed near their coupon. Two years is not a reason to abandon a long-term share plan; it is a reminder that the “winner” changes when the clock is short."
  },
  {
    n: 27, part: "D", marks: 0.5, key: "B",
    text: "What is the 'premium' on an insurance policy?",
    bullets: [],
    prompt: "",
    opts: {A:"The amount the insurance company pays you when you make a claim", B:"The regular amount you pay to keep your insurance active", C:"A bonus paid to you when the policy ends", D:"The part of the claim that you pay yourself"},
    why: "The premium is what you pay, usually every month or year, to keep cover in force. What the company pays you after a valid claim is the claim amount, not the premium. A maturity bonus is something else. The part of a claim you pay yourself is the deductible or co-pay."
  },
];
// END FEEDBACK_BANK

var PARTS = "AAAAABBBBBBBCCCCCCCCDDDDDDD".split("");
var PART_MAX = { A: 2.5, B: 7, C: 12, D: 3.5 };
var NAVY = "#0f2438";
var CREAM = "#f7f5f0";
var GOLD = "#a89468";
var WHITE = "#fffcf7";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("FinLit")
    .addItem("Rebuild dashboard", "setup")
    .addItem("Repair shifted rows", "fixShiftedRowsNow")
    .addItem("Email student results", "emailStudentsFromSheetNow")
    .addItem("Fill pending answers", "fillPendingChoices_")
    .addToUi();
}

function doGet(e) {
  if (e && e.parameter && String(e.parameter.email || "").trim()) {
    return ContentService
      .createTextOutput(JSON.stringify({
        taken: emailAlreadyUsed_(e.parameter.email)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  setup();
  return ContentService
    .createTextOutput("FinLit dashboard rebuilt.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    ingestAttempt_(data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  migrateOldData_(ss);
  rebuildAll_();
}

function rebuildAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responses = getOrCreateSheet_(ss, "Responses");
  fillPendingChoices_();
  normalizeResponses_(responses);
  var data = readResponses_(responses);
  buildDashboard_(ss, data);
  buildQuestionAnalysis_(ss, data);
  buildDemographics_(ss, data);
  hideUnusedSheets_(ss);
}

function notify_(msg) {
  Logger.log(msg);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg).slice(0, 190), "FinLit", 15);
  } catch (e) {}
}

function responseUsedRow_(sheet) {
  var last = Number(sheet.getLastRow()) || 1;
  return last > 400 ? 400 : last;
}

function responseWidth_(sheet) {
  var width = Number(sheet.getLastColumn()) || 45;
  if (width < 45) width = 45;
  if (width > 60) width = 60;
  return width;
}

function isLockedStudent_(name) {
  var n = String(name || "").trim().toLowerCase();
  return n === "verushka patel" || n === "prathmesh khurana";
}

function pendingChoicePlans_() {
  return {
    "anonymous": {
      total: 20,
      parts: { A: 2.5, B: 7, C: 7.5, D: 3 },
      missed: "Q13 (A); Q14 (A); Q18 (B); Q23 (C)",
      answers: ["B","B","C","B","C","B","C","D","B","C","B","C","A","A","C","D","C","B","C","A","C","A","C","C","B","B","B"]
    },
    "atharv agarwal": {
      total: 22.5,
      parts: { A: 2.5, B: 5, C: 12, D: 3 },
      missed: "Q6 (A); Q7 (A); Q26 (A)",
      answers: ["B","B","C","B","C","A","A","D","B","C","B","C","C","B","C","D","C","C","C","A","C","A","B","C","B","A","B"]
    },
    "neil arya": {
      total: 22.5,
      parts: { A: 2.5, B: 7, C: 10.5, D: 2.5 },
      missed: "Q19 (B); Q22 (B); Q25 (D)",
      answers: ["B","B","C","B","C","B","C","D","B","C","B","C","C","B","C","D","C","C","B","A","C","B","B","C","D","B","B"]
    },
    "arin mathew": {
      total: 22.5,
      parts: { A: 2.5, B: 7, C: 10.5, D: 2.5 },
      missed: "Q13 (A); Q23 (C); Q25 (D)",
      answers: ["B","B","C","B","C","B","C","D","B","C","B","C","A","B","C","D","C","C","C","A","C","A","C","C","D","B","B"]
    }
  };
}

function fillPendingChoices_() {
  var sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Responses");
  ensureResponseHeaders_(sheet);
  var last = responseUsedRow_(sheet);
  if (last < 2) return;
  var width = responseWidth_(sheet);
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  var plans = pendingChoicePlans_();
  var changed = false;
  for (var r = 0; r < values.length; r++) {
    var name = String(values[r][1] || "").trim();
    if (!name || isLockedStudent_(name)) continue;
    var plan = plans[name.toLowerCase()];
    if (!plan) continue;
    var missed = String(values[r][17] || "").toLowerCase();
    var q1 = String(values[r][18] || "").trim().toLowerCase();
    var needs = missed.indexOf("to be updated") !== -1 || !q1 || q1 === "blank";
    if (!needs) continue;
    values[r][12] = plan.total;
    values[r][13] = plan.parts.A;
    values[r][14] = plan.parts.B;
    values[r][15] = plan.parts.C;
    values[r][16] = plan.parts.D;
    values[r][17] = plan.missed;
    for (var i = 0; i < 27; i++) values[r][18 + i] = plan.answers[i];
    changed = true;
  }
  if (changed) sheet.getRange(2, 1, values.length, width).setValues(values);
}

function isShiftedAttemptRow_(row) {
  var name = String(row[1] || "").trim().toLowerCase();
  if (isLockedStudent_(name)) return false;
  var email = String(row[2] || "").trim();
  var timed = String(row[11] || "").trim();
  var emailIsName = email && email.indexOf("@") < 0;
  var looksTime = /\d+(\.\d+)?\s*min/i.test(timed);
  var looksScore = /\d+(\.\d+)?\s*\/\s*25/.test(String(row[14] || "")) ||
    /\d+(\.\d+)?\s*\/\s*25/.test(String(row[13] || ""));
  var autoInScore = /^(yes|no)/i.test(String(row[12] || "").trim());
  var missedBlob = /picked|blank/i.test(String(row[15] || ""));
  return looksTime && autoInScore && (looksScore || missedBlob || emailIsName || name === "student");
}

function repairShiftedAttemptRow_(row) {
  var next = row.slice();
  while (next.length < 45) next.push("");
  var email = String(next[2] || "").trim();
  if (email && email.indexOf("@") < 0) {
    next[1] = email;
    next[2] = "";
  }
  next[5] = next[6];
  next[6] = next[7];
  next[7] = next[8];
  next[8] = next[9];
  next[9] = next[10];
  var timeRaw = String(next[11] || "");
  var timeMin = Number(timeRaw.replace(/ min/i, "").trim());
  next[10] = isFinite(timeMin) && timeRaw ? timeMin : next[10];
  next[11] = /yes/i.test(String(row[12] || "")) ? "Yes" : "No";
  var answers = [];
  for (var i = 0; i < 27; i++) answers.push(parseChoice_(next[16 + i], i));
  var scored = scoreChoices_(answers);
  next[12] = scored.total;
  next[13] = scored.parts.A;
  next[14] = scored.parts.B;
  next[15] = scored.parts.C;
  next[16] = scored.parts.D;
  next[17] = scored.missed;
  for (var j = 0; j < 27; j++) next[18 + j] = scored.choices[j];
  return next;
}

function repairShiftedAttemptRows_() {
  var sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Responses");
  ensureResponseHeaders_(sheet);
  var last = responseUsedRow_(sheet);
  if (last < 2) return { count: 0, names: [] };
  var width = responseWidth_(sheet);
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  var names = [];
  for (var r = 0; r < values.length; r++) {
    if (!rowHasContent_(values[r])) continue;
    if (isLockedStudent_(values[r][1])) continue;
    if (!isShiftedAttemptRow_(values[r])) continue;
    values[r] = repairShiftedAttemptRow_(values[r]);
    names.push(String(values[r][1] || "student").trim());
  }
  if (names.length) sheet.getRange(2, 1, values.length, width).setValues(values);
  compactEmptyResponseRows_(sheet);
  return { count: names.length, names: names };
}

function rowHasContent_(row) {
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] || "").trim()) return true;
  }
  return false;
}

function compactEmptyResponseRows_(sheet) {
  var last = responseUsedRow_(sheet);
  if (last < 2) return;
  var width = responseWidth_(sheet);
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  var kept = values.filter(rowHasContent_);
  if (kept.length === values.length) return;
  sheet.getRange(2, 1, last - 1, width).clearContent();
  if (kept.length) sheet.getRange(2, 1, kept.length, width).setValues(kept);
}

function fixShiftedRowsNow() {
  var result = repairShiftedAttemptRows_();
  SpreadsheetApp.flush();
  var msg = result.count
    ? "Fixed " + result.count + " row(s): " + result.names.join(", ")
    : "No shifted rows found. Open the Responses tab and check this script is bound to the FinLit spreadsheet.";
  notify_(msg);
}

function fixKiaraNow() {
  fixShiftedRowsNow();
}

function sendMail_(to, subject, textBody, htmlBody) {
  GmailApp.sendEmail(to, subject, textBody, {
    htmlBody: htmlBody || textBody,
    name: "The Skyward Project",
    replyTo: "verushkapatel4@gmail.com"
  });
}

function sendPreviewEmailToMe() {
  var choices = ["D","B","C","B","C","B","C","B","A","C","A","C","C","D","C","C","C","C","C","B","A","B","B","B","B","B","A"];
  var scored = scoreChoices_(choices);
  sendMail_(
    "verushkapatel4@gmail.com",
    "Your FinLit Index results — The Skyward Project",
    buildFeedbackText_("Jiya (preview)", scored),
    buildFeedbackEmail_("Jiya (preview)", choices, scored)
  );
  notify_("Preview sent to verushkapatel4@gmail.com. Check Inbox and Sent.");
}

function authorizeMailNow() {
  sendMail_(
    "verushkapatel4@gmail.com",
    "FinLit mail is allowed",
    "This is only a permission check. You can delete it."
  );
}

function ingestAttempt_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(ss, "Responses");
  ensureResponseHeaders_(sheet);
  data = normalizeIncoming_(data);
  if (!validEmail_(data.email)) {
    throw new Error("A real email address is required.");
  }
  if (emailAlreadyUsed_(data.email)) return;
  var scored = scoreAttempt_(data);
  appendResponseRow_(sheet, buildResponseRow_(data, scored));
  try {
    sendStudentFeedback_(data, scored);
  } catch (err) {
    console.error("Student feedback email failed: " + err);
  }
}

function validEmail_(raw) {
  var email = String(raw || "").trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function normalizeIncoming_(data) {
  data = data || {};
  var name = String(data.name || "").trim();
  var email = String(data.email || "").trim();
  if (email && email.indexOf("@") < 0 && (!name || /^student$/i.test(name))) {
    name = email;
    email = "";
  }
  data.name = name;
  data.email = email;
  data.bg_transport = data.bg_transport || data.transport || "";
  return data;
}

function appendResponseRow_(sheet, row) {
  var width = responseHeaders_().length;
  var next = row.slice();
  while (next.length < width) next.push("");
  if (next.length > width) next = next.slice(0, width);
  sheet.getRange(nextResponseRow_(sheet), 1, 1, width).setValues([next]);
}

function nextResponseRow_(sheet) {
  var last = Math.max(responseUsedRow_(sheet), 2);
  var width = responseWidth_(sheet);
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  for (var i = 0; i < values.length; i++) {
    if (!rowHasContent_(values[i])) return i + 2;
  }
  return last + 1;
}

function emailAlreadyUsed_(raw) {
  var email = String(raw || "").trim().toLowerCase();
  if (!validEmail_(email)) return false;
  var sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Responses");
  var last = responseUsedRow_(sheet);
  if (last < 2) return false;
  var vals = sheet.getRange(2, 3, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "").trim().toLowerCase() === email) return true;
  }
  return false;
}

function shouldEmailStudent_(data) {
  return validEmail_(data && data.email);
}

function sendStudentFeedback_(data, scored) {
  if (!shouldEmailStudent_(data)) return;
  var choices = (scored && scored.choices && scored.choices.length)
    ? scored.choices
    : extractChoices_(data);
  var name = String(data.name || "Student").trim() || "Student";
  sendMail_(
    String(data.email).trim(),
    "Your FinLit Index results — The Skyward Project",
    buildFeedbackText_(name, scored),
    buildFeedbackEmail_(name, choices, scored)
  );
}

function emailStudentsFromSheetNow() {
  var sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Responses");
  var last = responseUsedRow_(sheet);
  if (last < 2) {
    notify_("No response rows to email.");
    return;
  }
  var width = responseWidth_(sheet);
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  var sent = [];
  var skipped = [];
  for (var r = 0; r < values.length; r++) {
    if (!rowHasContent_(values[r])) continue;
    var name = String(values[r][1] || "").trim();
    if (isLockedStudent_(name) || /^anonymous$/i.test(name)) continue;
    var email = String(values[r][2] || "").trim();
    if (!validEmail_(email)) {
      skipped.push(name || ("row " + (r + 2)));
      continue;
    }
    var choices = [];
    for (var i = 0; i < 27; i++) choices.push(parseChoice_(values[r][18 + i], i));
    var scored = scoreChoices_(choices);
    var submitted = parseScore_(values[r][12]);
    if (submitted != null) scored.total = submitted;
    try {
      sendStudentFeedback_({ name: name, email: email }, scored);
      sent.push(name + " <" + email + ">");
    } catch (err) {
      skipped.push((name || email) + " (" + err + ")");
    }
  }
  notify_(
    (sent.length ? "Emailed " + sent.length + ": " + sent.join("; ") : "No emails sent.") +
    (skipped.length ? " Skipped: " + skipped.join("; ") : "")
  );
}

function buildFeedbackText_(name, scored) {
  return "Hello " + name + ",\n\nThank you for sitting the FinLit Index. Your score is " +
    scored.total + " / 25.\nPart A: " + scored.parts.A + " / 2.5 · Part B: " +
    scored.parts.B + " / 7 · Part C: " + scored.parts.C + " / 12 · Part D: " +
    scored.parts.D + " / 3.5\n\nOpen this message as HTML to see the questions you missed and why the correct option holds up.\n\nThe Skyward Project · Pune & Mumbai 2026";
}

function buildFeedbackEmail_(name, choices, scored) {
  var missed = [];
  var i, q, pick, optLine, bullets, letter;
  for (i = 0; i < FEEDBACK_BANK.length; i++) {
    q = FEEDBACK_BANK[i];
    pick = String(choices[i] || "").trim();
    if (/^blank$/i.test(pick)) pick = "";
    if (pick === q.key) continue;
    bullets = (q.bullets || []).map(function (b) {
      return "<li style=\"margin:0 0 4px;\">" + escHtml_(b) + "</li>";
    }).join("");
    optLine = "";
    for (letter = 0; letter < 4; letter++) {
      var L = ["A", "B", "C", "D"][letter];
      var isKey = L === q.key;
      optLine += "<div style=\"margin:0 0 6px;padding:8px 10px;background:" +
        (isKey ? "#f7f5f0" : "transparent") + ";border:1px solid " +
        (isKey ? "#a89468" : "#d9d5cc") + ";color:#0f2438;\">" +
        "<span style=\"font-family:Arial,Helvetica,sans-serif;font-size:11px;letter-spacing:.08em;color:#a89468;\">" + L + "</span> " +
        escHtml_(q.opts[L]) +
        (isKey ? " <span style=\"font-family:Arial,Helvetica,sans-serif;font-size:11px;letter-spacing:.08em;color:#a89468;\">CORRECT</span>" : "") +
        "</div>";
    }
    missed.push(
      "<div style=\"border-top:1px solid #d9d5cc;padding:20px 0;\">" +
      "<p style=\"margin:0 0 8px;font-family:Arial,Helvetica,sans-serif;font-size:11px;letter-spacing:.16em;text-transform:uppercase;color:#a89468;\">Question " + q.n + " · Part " + q.part + "</p>" +
      "<p style=\"margin:0 0 8px;color:#2b2b2b;\">" + escHtml_(q.text) + "</p>" +
      (bullets ? "<ul style=\"margin:0 0 10px 18px;padding:0;color:#2b2b2b;\">" + bullets + "</ul>" : "") +
      (q.prompt ? "<p style=\"margin:0 0 12px;color:#0f2438;font-style:italic;\">" + escHtml_(q.prompt) + "</p>" : "") +
      optLine +
      "<p style=\"margin:14px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#5c574f;\">You chose: <strong style=\"color:#0f2438;\">" + (pick ? pick : "blank") + "</strong></p>" +
      "<p style=\"margin:10px 0 0;color:#0f2438;\"><strong>Why " + q.key + " is right.</strong> " + escHtml_(q.why) + "</p>" +
      "</div>"
    );
  }
  var bodyMissed = missed.length
    ? missed.join("")
    : "<p style=\"color:#0f2438;\">You did not miss any questions.</p>";
  return (
    "<div style=\"margin:0;padding:28px 16px;background:#f7f5f0;color:#2b2b2b;\">" +
    "<div style=\"max-width:620px;margin:0 auto;background:#fffcf7;border:1px solid #d9d5cc;padding:32px 28px;font-family:Georgia,'Times New Roman',Times,serif;line-height:1.6;color:#2b2b2b;\">" +
    "<p style=\"margin:0 0 10px;text-align:center;font-family:Georgia,'Times New Roman',Times,serif;font-size:22px;color:#0f2438;letter-spacing:.02em;\"><em>The</em> Skyward Project</p>" +
    "<div style=\"height:4px;border-top:1px solid #0f2438;border-bottom:1px solid #0f2438;margin:0 0 18px;\"></div>" +
    "<p style=\"margin:0 0 18px;text-align:center;font-family:Arial,Helvetica,sans-serif;font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:#a89468;\">FinLit Index · Results</p>" +
    "<h1 style=\"margin:0 0 14px;font-size:26px;line-height:1.2;color:#0f2438;font-weight:normal;\">Hello " + escHtml_(name) + "</h1>" +
    "<p style=\"margin:0 0 16px;\">Thank you for sitting the paper. Your mark is not shown on the website. It is here, in this edition, along with the questions to look at again.</p>" +
    "<div style=\"background:#f7f5f0;border:1px solid #d9d5cc;padding:16px 18px;margin:0 0 22px;\">" +
    "<p style=\"margin:0 0 6px;font-family:Arial,Helvetica,sans-serif;font-size:11px;letter-spacing:.16em;text-transform:uppercase;color:#a89468;\">Your score</p>" +
    "<p style=\"margin:0;font-size:32px;line-height:1.1;color:#0f2438;\">" + scored.total + " <span style=\"font-size:16px;color:#5c574f;\">/ 25</span></p>" +
    "<p style=\"margin:10px 0 0;font-size:14px;color:#5c574f;\">Part A " + scored.parts.A + " / 2.5 &nbsp;·&nbsp; Part B " + scored.parts.B + " / 7 &nbsp;·&nbsp; Part C " + scored.parts.C + " / 12 &nbsp;·&nbsp; Part D " + scored.parts.D + " / 3.5</p>" +
    "</div>" +
    "<p style=\"margin:0 0 8px;font-size:20px;color:#0f2438;\">What to work on</p>" +
    "<p style=\"margin:0 0 8px;\">Below are the questions you missed, the option that holds up, and why. If a question felt hard, that is useful information — not a verdict on you.</p>" +
    bodyMissed +
    "<div style=\"height:4px;border-top:1px solid #0f2438;border-bottom:1px solid #0f2438;margin:24px 0 14px;\"></div>" +
    "<p style=\"margin:0;font-size:13px;color:#5c574f;\">This test measures how you reason about money as a general skill. Real choices still depend on your family, income, and goals.</p>" +
    "<p style=\"margin:8px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#a89468;\">The Skyward Project · Pune &amp; Mumbai 2026</p>" +
    "</div></div>"
  );
}

function escHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function scoreAttempt_(data) {
  var choices = extractChoices_(data);
  return applySubmittedScore_(data, scoreChoices_(choices));
}

function extractChoices_(data) {
  var choices = [];
  for (var i = 0; i < 27; i++) {
    var n = i + 1;
    choices.push(parseChoice_(data["q" + n + "_choice"] || data["q" + n], i));
  }
  return choices;
}

function parseChoice_(raw, index) {
  raw = String(raw || "").trim();
  if (!raw || /^blank$/i.test(raw) || raw === "OK") return "";
  if (/^[ABCD]$/i.test(raw)) return raw.toUpperCase();
  var marked = raw.match(/X\s*\(\s*([ABCD])\s*\)/i) || raw.match(/picked\s+([ABCD])/i);
  return marked ? marked[1].toUpperCase() : "";
}

function scoreChoices_(choices) {
  var parts = { A: 0, B: 0, C: 0, D: 0 };
  var total = 0;
  var missed = [];
  for (var i = 0; i < 27; i++) {
    var choice = choices[i] || "";
    if (choice && choice === KEYS[i]) {
      parts[PARTS[i]] += MARKS[i];
      total += MARKS[i];
    } else {
      missed.push("Q" + (i + 1) + (choice ? " (" + choice + ")" : " (blank)"));
    }
  }
  return {
    choices: choices.map(function (c) { return c || "blank"; }),
    parts: parts,
    total: Math.round(total * 100) / 100,
    missed: missed.length ? missed.join("; ") : "None"
  };
}

function responseHeaders_() {
  var headers = [
    "Timestamp", "Name", "Email", "Grade", "School",
    "Transport", "Parents occupation", "Area", "Devices", "Income",
    "Time used (min)", "Timed out", "Score / 25",
    "Part A / 2.5", "Part B / 7", "Part C / 12", "Part D / 3.5",
    "Missed questions"
  ];
  for (var i = 1; i <= 27; i++) headers.push("Q" + i);
  return headers;
}

function ensureResponseHeaders_(sheet) {
  var headers = responseHeaders_();
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(5);
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(18, 280);
}

function buildResponseRow_(data, scored) {
  var timeMin = data.time_used_min;
  if (timeMin === undefined || timeMin === null || timeMin === "") {
    timeMin = Number(String(data.time_used || "").replace(" min", "")) || "";
  }
  var timedOut = String(data.auto_submitted || "").toLowerCase().indexOf("yes") !== -1 ? "Yes" : "No";
  var row = [
    data.submitted_at || new Date().toISOString(),
    data.name || "",
    data.email || "",
    data.grade || "",
    canonicalSchool_(data.school || ""),
    data.bg_transport || data.transport || "",
    data.bg_parents || "",
    data.bg_area || "",
    data.bg_devices || "",
    data.bg_income ? canonicalIncome_(data.bg_income) : (data.income ? canonicalIncome_(data.income) : ""),
    timeMin,
    timedOut,
    scored.total,
    scored.parts.A,
    scored.parts.B,
    scored.parts.C,
    scored.parts.D,
    scored.missed
  ];
  return row.concat(scored.choices);
}

function migrateOldData_(ss) {
  var names = ss.getSheets().map(function (s) { return s.getName(); });
  var source = null;
  if (names.indexOf("Responses") !== -1 && ss.getSheetByName("Responses").getLastRow() > 1) {
    source = ss.getSheetByName("Responses");
  } else if (names.indexOf("Sheet1") !== -1 && ss.getSheetByName("Sheet1").getLastRow() > 1) {
    source = ss.getSheetByName("Sheet1");
  }
  if (!source) {
    ensureResponseHeaders_(getOrCreateSheet_(ss, "Responses"));
    return;
  }

  var values = source.getDataRange().getValues();
  if (values.length < 2) {
    ensureResponseHeaders_(getOrCreateSheet_(ss, "Responses"));
    return;
  }

  var headers = values[0].map(function (h) { return String(h); });
  if (!isLegacyDump_(headers)) {
    ensureResponseHeaders_(getOrCreateSheet_(ss, "Responses"));
    return;
  }

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var rec = {};
    headers.forEach(function (h, c) { rec[h] = values[r][c]; });
    if (!(rec.name || rec.Name || rec.q1 || rec.Q1 || rec.q1_choice)) continue;
    var choices = [];
    for (var i = 0; i < 27; i++) choices.push(choiceFromRecord_(rec, i));
    var scored = applySubmittedScore_({
      score: rec.score != null && rec.score !== "" ? rec.score : rec["Score / 25"],
      score_total: rec.score_total,
      part_a: rec.part_a != null ? rec.part_a : rec["Part A / 2.5"],
      part_b: rec.part_b != null ? rec.part_b : rec["Part B / 7"],
      part_c: rec.part_c != null ? rec.part_c : rec["Part C / 12"],
      part_d: rec.part_d != null ? rec.part_d : rec["Part D / 3.5"],
      wrong_questions: rec.wrong_questions || rec["Missed questions"]
    }, scoreChoices_(choices));
    rows.push(buildResponseRow_({
      submitted_at: rec["Timestamp"] || rec.submitted_at,
      name: rec.Name || rec.name,
      email: rec.Email || rec.email || "",
      grade: rec.Grade || rec.grade,
      school: rec.School || rec.school,
      bg_transport: rec.Transport || rec.bg_transport,
      bg_parents: rec["Parents occupation"] || rec.bg_parents,
      bg_area: rec.Area || rec.bg_area,
      bg_devices: rec.Devices || rec.bg_devices,
      bg_income: rec.Income || rec.bg_income,
      time_used_min: rec["Time used (min)"] || rec.time_used_min,
      time_used: rec.time_used,
      auto_submitted: rec["Timed out"] === "Yes" ? "yes" : rec.auto_submitted
    }, scored));
  }

  var dest = getOrCreateSheet_(ss, "Responses");
  dest.clear();
  ensureResponseHeaders_(dest);
  if (rows.length) dest.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function isLegacyDump_(headers) {
  var set = {};
  headers.forEach(function (h) { set[String(h)] = 1; });
  return !!(set.q1 || set.q1_choice || set.profession || set.correct_count ||
    set["Q1 correct"] || set["Correct / 27"] || set.score_total);
}

function choiceFromRecord_(rec, i) {
  var n = i + 1;
  return parseChoice_(
    rec["q" + n + "_choice"] || rec["q" + n] || rec["Q" + n],
    i
  );
}

function applySubmittedScore_(data, scored) {
  var submitted = parseScore_(data.score);
  if (submitted == null) submitted = parseScore_(data.score_total);
  if (submitted != null) scored.total = submitted;
  var a = parseScore_(data.part_a);
  var b = parseScore_(data.part_b);
  var c = parseScore_(data.part_c);
  var d = parseScore_(data.part_d);
  if (a != null) scored.parts.A = a;
  if (b != null) scored.parts.B = b;
  if (c != null) scored.parts.C = c;
  if (d != null) scored.parts.D = d;
  if (data.wrong_questions) scored.missed = data.wrong_questions;
  return scored;
}

function parseScore_(raw) {
  if (raw === "" || raw == null) return null;
  if (typeof raw === "number" && isFinite(raw)) return raw;
  var m = String(raw).replace(/,/g, "").match(/([0-9]+(?:\.[0-9]+)?)/);
  return m ? Number(m[1]) : null;
}

function normalizeResponses_(sheet) {
  ensureResponseHeaders_(sheet);
  var last = responseUsedRow_(sheet);
  if (last < 2) return;
  var width = responseWidth_(sheet);
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  var out = values.map(function (row) {
    var next = row.slice();
    next[4] = canonicalSchool_(next[4] || "");
    if (next[9]) next[9] = canonicalIncome_(next[9]);
    return next;
  });
  sheet.getRange(2, 1, out.length, width).setValues(out);
}

function readResponses_(sheet) {
  var last = responseUsedRow_(sheet);
  if (last < 2) return { rows: [], n: 0 };
  var values = sheet.getRange(2, 1, last - 1, responseWidth_(sheet)).getValues();
  var rows = values.filter(function (r) { return String(r[1]).trim() !== ""; });
  return { rows: rows, n: rows.length };
}

function styleTitle_(range, size) {
  range.setFontFamily("Georgia")
    .setFontWeight("bold")
    .setFontColor(NAVY)
    .setFontSize(size);
}

function buildDashboard_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Dashboard");
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(GOLD);
  sheet.getRange("A1:L55").setBackground(CREAM);
  [220, 100, 110, 24, 200, 90, 90, 24, 160, 90].forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  styleTitle_(sheet.getRange("A1"), 22);
  sheet.getRange("A1").setValue("The Skyward Project  ·  FinLit Index");
  sheet.getRange("A2").setValue("Pune 2026  ·  scores out of 25").setFontColor(GOLD).setFontStyle("italic");

  var n = data.n;
  var scores = data.rows.map(function (r) { return parseScore_(r[12]) || 0; });
  var timedOut = data.rows.filter(function (r) {
    return String(r[11]).toLowerCase().indexOf("yes") !== -1;
  }).length;

  var kpis = [
    ["Attempts", n],
    ["Average score", n ? round1_(avg_(scores)) : 0],
    ["Median score", n ? round1_(median_(scores)) : 0],
    ["Highest", n ? round1_(Math.max.apply(null, scores)) : 0],
    ["Lowest", n ? round1_(Math.min.apply(null, scores)) : 0],
    ["Timed out", timedOut]
  ];
  sheet.getRange("A4").setValue("Overview").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(5, 1, kpis.length, 2).setValues(kpis);
  sheet.getRange(5, 1, kpis.length, 1).setFontColor("#5c574f");
  sheet.getRange(5, 2, kpis.length, 1).setFontWeight("bold").setFontSize(16).setFontColor(NAVY);

  var bands = [
    ["Score band", "Students"],
    ["0–5", countBand_(scores, 0, 5)],
    ["5–10", countBand_(scores, 5, 10, true)],
    ["10–15", countBand_(scores, 10, 15, true)],
    ["15–20", countBand_(scores, 15, 20, true)],
    ["20–25", countBand_(scores, 20, 25)]
  ];
  sheet.getRange("A13").setValue("Score distribution").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(14, 1, bands.length, 2).setValues(bands);
  paintHeader_(sheet.getRange(14, 1, 1, 2));

  var gradeTable = [["Grade", "Students", "Average"]];
  var gradeMap = groupAvg_(data.rows, 3, 12);
  ["9", "10", "11", "12"].forEach(function (g) {
    var item = gradeMap[g] || { count: 0, avg: 0 };
    gradeTable.push([g, item.count, item.count ? round1_(item.avg) : 0]);
  });
  sheet.getRange("A22").setValue("By grade").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(23, 1, gradeTable.length, 3).setValues(gradeTable);
  paintHeader_(sheet.getRange(23, 1, 1, 3));

  var partAvgs = [
    ["Part", "Average", "Maximum"],
    ["A  Warm up", n ? round1_(avg_(col_(data.rows, 13))) : 0, PART_MAX.A],
    ["B  Everyday", n ? round1_(avg_(col_(data.rows, 14))) : 0, PART_MAX.B],
    ["C  Read carefully", n ? round1_(avg_(col_(data.rows, 15))) : 0, PART_MAX.C],
    ["D  Market awareness", n ? round1_(avg_(col_(data.rows, 16))) : 0, PART_MAX.D]
  ];
  sheet.getRange("A30").setValue("Part averages").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(31, 1, partAvgs.length, 3).setValues(partAvgs);
  paintHeader_(sheet.getRange(31, 1, 1, 3));

  var incomeTable = incomeTable_(data.rows);
  sheet.getRange("A38").setValue("Income and performance").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(39, 1, incomeTable.length, 3).setValues(incomeTable);
  paintHeader_(sheet.getRange(39, 1, 1, 3));

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange("A14:B19"), 4, 5, "Students in each score band");
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange("A23:C27"), 18, 5, "Average score by grade");
  addChart_(sheet, Charts.ChartType.BAR, sheet.getRange("A31:B35"), 31, 5, "Average marks by part");
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange(39, 1, incomeTable.length, 2), 38, 9, "Average score by family income");
}

function buildQuestionAnalysis_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Question Analysis");
  sheet.clear();
  sheet.setTabColor(NAVY);
  sheet.setFrozenRows(1);
  var n = data.n;
  var table = [["Question", "Part", "Marks", "Correct answer", "Got it right", "Attempts", "% correct", "Most common pick"]];
  var chartRows = [["Question", "% correct"]];

  for (var i = 0; i < 27; i++) {
    var right = 0;
    var counts = { A: 0, B: 0, C: 0, D: 0, blank: 0 };
    data.rows.forEach(function (r) {
      var ch = parseChoice_(r[18 + i], i) || "blank";
      counts[ch] = (counts[ch] || 0) + 1;
      if (ch === KEYS[i]) right++;
    });
    var top = "—", topN = -1;
    Object.keys(counts).forEach(function (k) {
      if (counts[k] > topN) { topN = counts[k]; top = k; }
    });
    var pct = n ? Math.round((right / n) * 1000) / 10 : 0;
    table.push(["Q" + (i + 1), PARTS[i], MARKS[i], KEYS[i], right, n, pct, n ? top : "—"]);
    chartRows.push(["Q" + (i + 1), pct]);
  }

  sheet.getRange(1, 1, table.length, 8).setValues(table);
  paintHeader_(sheet.getRange(1, 1, 1, 8));
  sheet.getRange(1, 10, chartRows.length, 2).setValues(chartRows);
  sheet.hideColumns(10, 2);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(8, 160);
  sheet.getRange("A1:H1").setNote("Correct answers are for evaluators. They are not shown on the student website.");

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange(1, 10, chartRows.length, 2), 1, 9, "% of students correct on each question");
}

function buildDemographics_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Demographics");
  sheet.clear();
  sheet.setTabColor(GOLD);
  sheet.getRange("A1:P50").setBackground(CREAM);
  styleTitle_(sheet.getRange("A1"), 18);
  sheet.getRange("A1").setValue("Background patterns");
  sheet.getRange("A2").setValue("Refreshes after every submission and whenever you run Rebuild dashboard.").setFontColor("#5c574f");

  writeGroupTable_(sheet, 4, 1, "By school", groupAvg_(data.rows, 4, 12, canonicalSchool_), ["School", "Students", "Average score"]);
  writeGroupTable_(sheet, 4, 5, "By transport", groupAvg_(data.rows, 5, 12), ["Transport", "Students", "Average score"]);
  var incomeTable = incomeTable_(data.rows);
  sheet.getRange(4, 9).setValue("Income and performance").setFontWeight("bold").setFontColor(NAVY).setFontSize(12);
  sheet.getRange(5, 9, incomeTable.length, 3).setValues(incomeTable);
  paintHeader_(sheet.getRange(5, 9, 1, 3));
  writeGroupTable_(sheet, 20, 1, "By area", groupAvg_(data.rows, 7, 12), ["Area", "Students", "Average score"]);

  var deviceCounts = {};
  data.rows.forEach(function (r) {
    String(r[8] || "").split(",").map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (d) {
      deviceCounts[d] = (deviceCounts[d] || 0) + 1;
    });
  });
  var deviceTable = [["Device", "Students"]];
  Object.keys(deviceCounts).sort().forEach(function (k) { deviceTable.push([k, deviceCounts[k]]); });
  if (deviceTable.length === 1) deviceTable.push(["—", 0]);
  sheet.getRange(20, 5).setValue("Devices owned").setFontWeight("bold").setFontColor(NAVY);
  sheet.getRange(21, 5, deviceTable.length, 2).setValues(deviceTable);
  paintHeader_(sheet.getRange(21, 5, 1, 2));

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  var schoolN = Math.max(Object.keys(groupAvg_(data.rows, 4, 12, canonicalSchool_)).length, 1);
  var transportN = Math.max(Object.keys(groupAvg_(data.rows, 5, 12)).length, 1);
  addChart_(sheet, Charts.ChartType.BAR, sheet.getRange(5, 1, 1 + schoolN, 3), 4, 12, "Average score by school");
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange(5, 5, 1 + transportN, 2), 20, 12, "How students come to school");
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange(5, 9, incomeTable.length, 2), 32, 9, "Average score by family income");
}

function writeGroupTable_(sheet, row, col, title, grouped, headers) {
  sheet.getRange(row, col).setValue(title).setFontWeight("bold").setFontColor(NAVY).setFontSize(12);
  var keys = Object.keys(grouped).sort();
  var table = [headers];
  if (!keys.length) table.push(["—", 0, 0]);
  keys.forEach(function (k) {
    table.push([k || "—", grouped[k].count, round1_(grouped[k].avg)]);
  });
  sheet.getRange(row + 1, col, table.length, 3).setValues(table);
  paintHeader_(sheet.getRange(row + 1, col, 1, 3));
}

function paintHeader_(range) {
  range.setFontWeight("bold").setBackground(NAVY).setFontColor(WHITE);
}

function addChart_(sheet, type, range, row, col, title) {
  try {
    var chart = sheet.newChart()
      .setChartType(type)
      .addRange(range)
      .setOption("title", title)
      .setOption("legend", { position: "none" })
      .setOption("colors", [NAVY])
      .setOption("backgroundColor", WHITE)
      .setOption("hAxis", { textStyle: { color: NAVY } })
      .setOption("vAxis", { textStyle: { color: NAVY }, minValue: 0 })
      .setOption("titleTextStyle", { color: NAVY, fontSize: 12, bold: true })
      .setPosition(row, col, 0, 0)
      .build();
    sheet.insertChart(chart);
  } catch (err) {
    Logger.log("Chart skipped: " + title + " — " + err);
  }
}

function hideUnusedSheets_(ss) {
  var keep = { Responses: 1, Dashboard: 1, "Question Analysis": 1, Demographics: 1 };
  ss.getSheets().forEach(function (s) {
    if (!keep[s.getName()] && ss.getSheets().length > 1) {
      try { s.hideSheet(); } catch (e) {}
    }
  });
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function groupAvg_(rows, keyIndex, valueIndex, canonFn) {
  var map = {};
  rows.forEach(function (r) {
    var key = String(r[keyIndex] || "").trim() || "Not given";
    if (canonFn) key = canonFn(key) || "Not given";
    if (!map[key]) map[key] = { count: 0, sum: 0 };
    map[key].count++;
    map[key].sum += parseScore_(r[valueIndex]) || 0;
  });
  Object.keys(map).forEach(function (k) {
    map[k].avg = map[k].count ? map[k].sum / map[k].count : 0;
  });
  return map;
}

function canonicalSchool_(name) {
  var raw = String(name || "").trim();
  if (!raw || raw === "Not given") return raw;
  var n = raw.toLowerCase().replace(/&/g, "and").replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
  var compact = n.replace(/ /g, "");
  if (compact === "dais") return "DAIS";
  if (/\bdais\b/.test(n)) return "DAIS";
  if (n.indexOf("dhirubhai") !== -1 && n.indexOf("ambani") !== -1) return "DAIS";
  return raw;
}

function canonicalIncome_(label) {
  var s = String(label || "").trim();
  if (!s) return "Not given";
  var n = s.replace(/[–—]/g, "-").replace(/\s+/g, " ").trim();
  if (/0\s*[-to]+\s*25/.test(n)) return "₹0 to 25,000";
  if (/25,001\s*[-to]+\s*50/.test(n)) return "₹25,001 to 50,000";
  if (/50,001\s*[-to]+\s*1/.test(n)) return "₹50,001 to 1,00,000";
  if (/1,00,001\s*[-to]+\s*2/.test(n)) return "₹1,00,001 to 2,50,000";
  if (/2,50,001|above/i.test(n)) return "₹2,50,001 or above";
  if (/not sure/i.test(n)) return "Not sure";
  return s;
}

function incomeTable_(rows) {
  var map = groupAvg_(rows, 9, 12, canonicalIncome_);
  var order = [
    "₹0 to 25,000",
    "₹25,001 to 50,000",
    "₹50,001 to 1,00,000",
    "₹1,00,001 to 2,50,000",
    "₹2,50,001 or above",
    "Not sure"
  ];
  var table = [["Income", "Average score", "Students"]];
  var seen = {};
  order.forEach(function (k) {
    if (!map[k]) return;
    table.push([k, round1_(map[k].avg), map[k].count]);
    seen[k] = 1;
  });
  Object.keys(map).forEach(function (k) {
    if (!seen[k]) table.push([k, round1_(map[k].avg), map[k].count]);
  });
  if (table.length === 1) table.push(["—", 0, 0]);
  return table;
}

function col_(rows, i) {
  return rows.map(function (r) { return parseScore_(r[i]) || 0; });
}

function avg_(arr) {
  if (!arr.length) return 0;
  return arr.reduce(function (a, b) { return a + b; }, 0) / arr.length;
}

function median_(arr) {
  if (!arr.length) return 0;
  var a = arr.slice().sort(function (x, y) { return x - y; });
  var mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}

function countBand_(scores, from, to, exclusiveStart) {
  return scores.filter(function (s) {
    if (exclusiveStart) return s > from && s <= to;
    return s >= from && s <= to;
  }).length;
}
