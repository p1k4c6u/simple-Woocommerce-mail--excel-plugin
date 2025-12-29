const SHEET_NAME = "Orders";

const TOKEN = "t0ots1kene66b5f3";

const VERIFY_SIGNATURE = true; // false/true
const WEBHOOK_SECRET = "t0ots1kene66b5f3";

function doPost(e) {
  try {
    // tokeni checki, tolli amet
    const token = (e.parameter && e.parameter.token) ? String(e.parameter.token): "";
    if (token !== TOKEN) {
        return _json(401, { ok: false, error: "Unauthorized (tolli amet ei lase labi)" });
    }

    const raw = e.postData && e.postData.contents ? e.postData.contents : "";
    if (!raw) return _json(400, { ok: false, error: "empty body (puudu keha, kuhu matsid?)" });

    if (VERIFY_SIGNATURE) {
        const sig = _getHeader(e, "X-WC-Webhook-Signature");
        if (!sig) return _json(401, { ok: false, error: "Unauthorized (puudu allkiri)" });
        if (!_verifySignature(raw, WEBHOOK_SECRET, sig)) {
            return _json(401, { ok: false, error: "Unauthorized (vale allkiri)" });
        }
    }
    //detailid
    const order = JSON.parse(raw);

    const orderId = order.id ?? "";
    const total = order.total ?? "";

    const billingFirst = order.billing?.first_name ?? "";
    const billingLast = order.billing?.last_name ?? "";
    const billingPhone = order.billing?.phone ?? "";

    const shipFirst = order.shipping?.first_name ?? "";
    const shipLast = order.shipping?.last_name ?? "";
    const shipName  = [shipFirst, shipLast].filter(Boolean).join(" ").trim();

    const shipPhone = order.shipping?.phone ?? billingPhone;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return _json(500, { ok: false, error: `Sheet "${SHEET_NAME}" not found (tolli amet otsis lehte, mida pole olemas)` });

    sheet.appendRow([
        new Date(),
        orderId,
        total,
        shipName,
        shipPhone
    ]); 

    return _json(200, { ok: true, order_id: orderId });
    } catch (err) {
    return _json(500, { ok: false, error: String(err) });
    }
}

function _json(code, obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _getHeader(e, name) {
  const headers = e?.headers || e?.parameter || {};
  // Try a few common variants
  return headers[name] || headers[name.toLowerCase()] || headers[name.replace(/-/g, "_")] || "";
}