# 📬 Outlook Auto-Reply Macro: Order Confirmation Generator

This Outlook VBA macro listens for new order emails from WooCommerce and automatically drafts a professional reply email with:
- Extracted order ID
- A table of purchased products
- A placeholder for tracking code

> 🔁 The reply is saved as a draft so you can review or edit it before sending.

---

## ✨ Features

- Auto-triggers on new incoming emails with subject format:  
  `[OGAWA]: New order #xxxxx`
- Extracts product name, quantity, and color from WooCommerce HTML tables
- Cleans HTML to extract meaningful values
- Drafts a reply email with tracking placeholder and formatted product table

---

## ⚠️ Compatibility Note

This macro is designed to work with a specific **WooCommerce email template** layout (commonly used in `customer-completed-order` or `new order` notifications).

If you are using a **customized** or third-party WooCommerce email template:
- The macro **may fail to detect** or parse product rows correctly
- Table structure or tags (like `<td>`, `<tr>`, etc.) might differ
- Additional adjustments to the `GetProductTableRows` or `ExtractCell` functions may be required

> 💡 Always test with your live store’s actual email template before deployment.

---

## 📋 Sample Email Flow

1. A customer places an order via WooCommerce
2. Outlook receives an email with subject like:  
   `[OGAWA]: New order #2516`
3. The macro auto-parses the email and drafts a reply like:

Dear Customer,

Thank you for your purchase!
Your order #2516 is pending courier collection.

[Product table here]

You may track the order via https://www.jtexpress.sg/trackmyparcel
with tracking code: [INSERT_TRACKING_CODE]

You shall receive the order in 1–3 days after successful pickup.

Warm regards,
OGAWA Team

## 🛠️ Setup Instructions

1. Open **Outlook**
2. Press `ALT + F11` to open the **VBA Editor**
3. Go to `File > Import File…`
4. Select the `outlook-auto-draft-reply.bas` file
5. Save and close the editor

✅ Now your macro will trigger automatically on new order emails.

---

## 🔐 Security Note

- This macro is intended for internal customer support use
- Always **review the draft** email before sending manually
- Make sure macro security settings allow trusted code execution

---
