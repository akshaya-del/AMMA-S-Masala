/**
 * Amma's Home Masala & Pickles — Backend (Google Apps Script)
 * Connects the shop and the admin page through one Google Sheet.
 *
 * SETUP (one time, ~5 minutes):
 * 1. Create a new Google Sheet (e.g. "Amma's Shop").
 * 2. In the sheet: Extensions → Apps Script. Delete the sample code and paste this whole file.
 * 3. Change ADMIN_KEY below to your own strong secret password (only the admin should know it).
 * 4. At the top, choose the function "setup" and click ▶ Run. Allow the permissions.
 *    (This creates the sheet tabs. "Google hasn't verified this app" → Advanced → Go to project.)
 * 5. Click Deploy → New deployment → gear icon → "Web app".
 *      Execute as:      Me
 *      Who has access:  Anyone        <-- must be "Anyone", not "Anyone with Google account"
 *    Click Deploy and copy the "Web app URL" (it ends with /exec).
 * 6. Paste that URL into API_URL (in the SETTINGS section) of your website HTML file.
 *
 * If you edit this script later, use Deploy → Manage deployments → Edit → Version: New version
 * so the same URL keeps working.
 *
 * Sheet tabs created automatically:
 *   Orders, Events (shop activity), Feedback, Products, Settings, Reviews
 *
 * Security: customers can only ADD orders / feedback / activity / reviews (reviews wait for
 * admin approval) and READ visible products and approved reviews.
 * Everything else (viewing orders & customers, editing products, settings, order status)
 * needs ADMIN_KEY. After 10 wrong passwords, admin login is blocked for 15 minutes.
 */

const ADMIN_KEY = "change-this-password";   // <-- CHANGE THIS

const ORDER_HEADERS = ["Order ID", "Date", "Name", "Phone", "Address", "Pincode", "Area",
                       "Payment", "Items", "Items JSON", "Subtotal", "Status", "Notes"];
const EVENT_HEADERS = ["Date", "Session", "Type", "Product"];
const FEEDBACK_HEADERS = ["Date", "Type", "Message", "Order ID"];
const PRODUCT_HEADERS = ["ID", "Name", "Category", "Tags", "Price 250g", "Custom Prices JSON", "Spice",
                         "Shelf Life", "Ingredients", "Description", "Image", "Icon", "Color", "Visible", "In Stock"];
const SETTINGS_HEADERS = ["Key", "Value"];
const REVIEW_HEADERS = ["ID", "Date", "Name", "City", "Rating", "Product", "Review", "Status"];
const REVIEW_STATUSES = ["Pending", "Approved", "Hidden"];

const STATUSES = ["Pending", "Confirmed", "Packed", "Dispatched", "Delivered", "Cancelled"];
const EVENT_TYPES = ["visit", "view", "add", "checkout", "order"];
const SETTING_KEYS = ["heroTitle", "offerText"];
const EVENT_DAYS_RETURNED = 90;

// ---------------- Entry points ----------------
// Run this once from the Apps Script editor (select "setup" → Run) to create all sheet tabs.
function setup() {
  getSheet("Orders", ORDER_HEADERS);
  getSheet("Events", EVENT_HEADERS);
  getSheet("Feedback", FEEDBACK_HEADERS);
  getSheet("Products", PRODUCT_HEADERS);
  getSheet("Settings", SETTINGS_HEADERS);
  getSheet("Reviews", REVIEW_HEADERS);
  Logger.log("Setup done. Now Deploy → New deployment → Web app.");
}

function doGet(e) {
  try {
    return handleGet(e);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function handleGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === "ping") return json({ ok: true, message: "Connected" });
  if (action === "products") {
    // Public: only visible products, for the shop page
    return json({ ok: true, products: listProducts().filter(function (p) { return p.visible; }),
                  settings: getSettings(), reviews: listReviews(true) });
  }
  return ContentService.createTextOutput("Amma's Masala API is running.");
}

function doPost(e) {
  try {
    return handlePost(e);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function handlePost(e) {
  let body;
  try {
    if (!e || !e.postData) return json({ ok: false, error: "No data sent" });
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: "Bad request" });
  }

  // ---- Public (called by the shop) ----
  if (body.action === "create")   return json(createOrder(body.order || {}));
  if (body.action === "event")    return json(logEvent(body));
  if (body.action === "feedback") return json(saveFeedback(body));
  if (body.action === "review")   return json(submitReview(body.review || {}));

  // ---- Admin only ----
  const authError = checkKey(body.key);
  if (authError) return json({ ok: false, error: authError });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (body.action) {
      case "list":
        return json({ ok: true, orders: listOrders(), events: listEvents(), feedback: listFeedback(),
                      products: listProducts(), settings: getSettings(), reviews: listReviews(false) });
      case "update":       return json(updateOrder(body.id, body.status, body.notes));
      case "saveProducts": return json(saveProducts(body.products));
      case "saveSettings": return json(saveSettings(body.settings || {}));
      case "addReview":       return json(addReview(body.review || {}));
      case "setReviewStatus": return json(setReviewStatus(body.id, body.status));
      case "deleteReview":    return json(deleteReview(body.id));
      default:             return json({ ok: false, error: "Unknown action" });
    }
  } finally {
    lock.releaseLock();
  }
}

function checkKey(key) {
  if (ADMIN_KEY === "change-this-password") return "Set your own ADMIN_KEY in Apps Script first";
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("adminFails") || 0);
  if (fails >= 10) return "Too many wrong attempts. Try again in 15 minutes.";
  if (key !== ADMIN_KEY) {
    cache.put("adminFails", String(fails + 1), 900);
    return "Wrong admin key";
  }
  return null;
}

// ---------------- Helpers ----------------
function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Open this script from your Google Sheet (Extensions → Apps Script)");
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sh;
}

// Trim length and stop spreadsheet formula injection (values starting with = + - @)
function clean(v, max) {
  let s = String(v == null ? "" : v).slice(0, max || 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function unq(v) { return String(v == null ? "" : v).replace(/^'/, ""); }
function iso(d) { return d instanceof Date ? d.toISOString() : String(d); }

// ---------------- Orders ----------------
function createOrder(o) {
  const items = Array.isArray(o.items) ? o.items.slice(0, 100).map(function (i) {
    return {
      name: String(i.name || "").slice(0, 100),
      weight: String(i.weight || "").slice(0, 20),
      qty: Math.max(0, parseInt(i.qty, 10) || 0),
      priceUnit: Math.max(0, Number(i.priceUnit) || 0)
    };
  }) : [];
  if (!o.id || !o.name || !o.phone || items.length === 0) {
    return { ok: false, error: "Missing order details" };
  }
  const subtotal = items.reduce(function (s, i) { return s + i.qty * i.priceUnit; }, 0);
  const itemsText = items.map(function (i) {
    return i.name + " (" + i.weight + ") x" + i.qty + " = ₹" + (i.qty * i.priceUnit);
  }).join("\n");

  getSheet("Orders", ORDER_HEADERS).appendRow([
    clean(o.id, 30), new Date(), clean(o.name, 100), clean(o.phone, 20), clean(o.address, 500),
    clean(o.pincode, 10), clean(o.location, 50), clean(o.payment, 20),
    clean(itemsText, 5000), JSON.stringify(items), subtotal, "Pending", ""
  ]);
  return { ok: true };
}

function listOrders() {
  const rows = getSheet("Orders", ORDER_HEADERS).getDataRange().getValues().slice(1);
  return rows.filter(function (r) { return r[0]; }).map(function (r) {
    let items = [];
    try { items = JSON.parse(r[9] || "[]"); } catch (e) {}
    return {
      id: unq(r[0]), date: iso(r[1]), name: unq(r[2]), phone: unq(r[3]), address: unq(r[4]),
      pincode: unq(r[5]), location: unq(r[6]), payment: unq(r[7]), items: items,
      subtotal: Number(r[10]) || 0, status: String(r[11] || "Pending"), notes: unq(r[12])
    };
  });
}

function updateOrder(id, status, notes) {
  if (status && STATUSES.indexOf(status) === -1) return { ok: false, error: "Invalid status" };
  const sh = getSheet("Orders", ORDER_HEADERS);
  if (sh.getLastRow() < 2) return { ok: false, error: "Order not found" };
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (unq(ids[i][0]) === String(id)) {
      if (status) sh.getRange(i + 2, 12).setValue(status);
      if (notes !== undefined) sh.getRange(i + 2, 13).setValue(clean(notes, 1000));
      return { ok: true };
    }
  }
  return { ok: false, error: "Order not found" };
}

// ---------------- Events (shop activity) ----------------
function logEvent(b) {
  if (EVENT_TYPES.indexOf(b.type) === -1) return { ok: false, error: "Bad event" };
  getSheet("Events", EVENT_HEADERS).appendRow([
    new Date(), clean(b.session, 40), b.type, clean(b.product, 100)
  ]);
  return { ok: true };
}

function listEvents() {
  const cutoff = Date.now() - EVENT_DAYS_RETURNED * 86400000;
  const rows = getSheet("Events", EVENT_HEADERS).getDataRange().getValues().slice(1);
  const out = [];
  rows.forEach(function (r) {
    const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
    if (d.getTime() >= cutoff) out.push({ date: d.toISOString(), session: unq(r[1]), type: String(r[2]), product: unq(r[3]) });
  });
  return out;
}

// ---------------- Feedback ----------------
function saveFeedback(b) {
  const msg = String(b.message || "").trim();
  if (!msg) return { ok: false, error: "Empty message" };
  getSheet("Feedback", FEEDBACK_HEADERS).appendRow([
    new Date(), clean(b.type, 60), clean(msg, 2000), clean(b.orderId, 30)
  ]);
  return { ok: true };
}

function listFeedback() {
  const rows = getSheet("Feedback", FEEDBACK_HEADERS).getDataRange().getValues().slice(1);
  return rows.filter(function (r) { return r[2]; }).map(function (r) {
    return { date: iso(r[0]), type: unq(r[1]), message: unq(r[2]), orderId: unq(r[3]) };
  });
}

// ---------------- Products ----------------
function listProducts() {
  const rows = getSheet("Products", PRODUCT_HEADERS).getDataRange().getValues().slice(1);
  return rows.filter(function (r) { return r[1]; }).map(function (r) {
    let prices = {};
    try { prices = JSON.parse(r[5] || "{}"); } catch (e) {}
    return {
      id: Number(r[0]), name: unq(r[1]), cat: unq(r[2]),
      tags: unq(r[3]) ? unq(r[3]).split(",").map(function (t) { return t.trim(); }).filter(String) : [],
      price250: Number(r[4]) || 0, prices: prices, spice: unq(r[6]), shelf: unq(r[7]),
      ingredients: unq(r[8]), desc: unq(r[9]), img: unq(r[10]), icon: unq(r[11]), color: unq(r[12]),
      visible: r[13] !== false && String(r[13]).toUpperCase() !== "FALSE",
      inStock: r[14] !== false && String(r[14]).toUpperCase() !== "FALSE"
    };
  });
}

function saveProducts(list) {
  if (!Array.isArray(list)) return { ok: false, error: "Bad product list" };
  if (list.length > 500) return { ok: false, error: "Too many products" };
  const rows = list.map(function (p) {
    const prices = {};
    if (p.prices) Object.keys(p.prices).slice(0, 10).forEach(function (w) {
      const v = Number(p.prices[w]);
      if (v > 0) prices[String(w).slice(0, 10)] = v;
    });
    return [
      Number(p.id) || 0, clean(p.name, 100), clean(p.cat, 30),
      clean((Array.isArray(p.tags) ? p.tags : []).join(","), 100),
      Math.max(0, Number(p.price250) || 0), JSON.stringify(prices), clean(p.spice, 20),
      clean(p.shelf, 60), clean(p.ingredients, 500), clean(p.desc, 1000), clean(p.img, 500),
      clean(p.icon, 10), clean(p.color, 20), p.visible !== false, p.inStock !== false
    ];
  });
  const sh = getSheet("Products", PRODUCT_HEADERS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, PRODUCT_HEADERS.length).setValues(rows);
  return { ok: true };
}

// ---------------- Settings ----------------
function getSettings() {
  const out = {};
  getSheet("Settings", SETTINGS_HEADERS).getDataRange().getValues().slice(1).forEach(function (r) {
    if (SETTING_KEYS.indexOf(String(r[0])) > -1) out[r[0]] = unq(r[1]);
  });
  return out;
}

function saveSettings(s) {
  const sh = getSheet("Settings", SETTINGS_HEADERS);
  const rows = SETTING_KEYS.map(function (k) { return [k, clean(s[k], 300)]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return { ok: true };
}

// ---------------- Reviews ----------------
// Customers submit reviews from the website (status Pending); the admin approves them.
function submitReview(r) {
  const rating = Math.round(Number(r.rating));
  if (!String(r.name || "").trim() || !String(r.text || "").trim() || !(rating >= 1 && rating <= 5)) {
    return { ok: false, error: "Missing review details" };
  }
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get("reviewCount") || 0);
  if (count >= 50) return { ok: false, error: "Too many reviews right now. Please try again later." };
  cache.put("reviewCount", String(count + 1), 3600);
  appendReview(r, "Pending");
  return { ok: true };
}

// Admin adds a real review a customer sent (e.g. on WhatsApp) — shown straight away
function addReview(r) {
  if (!String(r.name || "").trim() || !String(r.text || "").trim()) {
    return { ok: false, error: "Customer name and review text are needed" };
  }
  return { ok: true, id: appendReview(r, "Approved") };
}

function appendReview(r, status) {
  const id = "R" + Date.now().toString(36) + Math.floor(Math.random() * 1000);
  const rating = Math.min(5, Math.max(1, Math.round(Number(r.rating)) || 5));
  getSheet("Reviews", REVIEW_HEADERS).appendRow([
    id, new Date(), clean(r.name, 60), clean(r.city, 40), rating, clean(r.product, 100), clean(r.text, 1000), status
  ]);
  return id;
}

function listReviews(approvedOnly) {
  const rows = getSheet("Reviews", REVIEW_HEADERS).getDataRange().getValues().slice(1);
  const all = rows.filter(function (r) { return r[0]; }).map(function (r) {
    return {
      id: unq(r[0]), date: iso(r[1]), name: unq(r[2]), city: unq(r[3]), rating: Number(r[4]) || 5,
      product: unq(r[5]), text: unq(r[6]), status: String(r[7] || "Pending")
    };
  }).reverse();
  return approvedOnly ? all.filter(function (r) { return r.status === "Approved"; }).slice(0, 50) : all;
}

function findReviewRow(id) {
  const sh = getSheet("Reviews", REVIEW_HEADERS);
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (unq(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function setReviewStatus(id, status) {
  if (REVIEW_STATUSES.indexOf(status) === -1) return { ok: false, error: "Invalid status" };
  const row = findReviewRow(id);
  if (!row) return { ok: false, error: "Review not found" };
  getSheet("Reviews", REVIEW_HEADERS).getRange(row, 8).setValue(status);
  return { ok: true };
}

function deleteReview(id) {
  const row = findReviewRow(id);
  if (!row) return { ok: false, error: "Review not found" };
  getSheet("Reviews", REVIEW_HEADERS).deleteRow(row);
  return { ok: true };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
