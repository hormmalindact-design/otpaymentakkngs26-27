/**
 * Netlify Function: secure server-side proxy to Google Apps Script.
 *
 * Required environment variables:
 *   APPS_SCRIPT_URL       = deployed Apps Script Web App URL ending in /exec
 *   API_GATEWAY_SECRET    = same value as Apps Script Script Property
 */
exports.handler = async function (event) {
  if (event.httpMethod === "OPTIONS") {
    return {
      statusCode: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Allow-Methods": "POST, OPTIONS"
      },
      body: ""
    };
  }

  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({ success: false, error: "Method not allowed" })
    };
  }

  const appsScriptUrl = process.env.APPS_SCRIPT_URL;
  const gatewaySecret = process.env.API_GATEWAY_SECRET;

  if (!appsScriptUrl || !gatewaySecret) {
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({
        success: false,
        error: "Netlify environment variables are not configured."
      })
    };
  }

  let requestData;
  try {
    requestData = JSON.parse(event.body || "{}");
  } catch (error) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({ success: false, error: "Invalid JSON request." })
    };
  }

  const allowedActions = new Set(["getTeachersList", "getPaymentData", "chatWithAI"]);
  if (!allowedActions.has(String(requestData.action || ""))) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({ success: false, error: "Unsupported API action." })
    };
  }

  try {
    const upstream = await fetch(appsScriptUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: requestData.action,
        payload: requestData.payload || {},
        gatewaySecret
      }),
      redirect: "follow"
    });

    const bodyText = await upstream.text();
    let data;
    try {
      data = JSON.parse(bodyText);
    } catch (error) {
      return {
        statusCode: 502,
        headers: { "Content-Type": "application/json; charset=utf-8" },
        body: JSON.stringify({
          success: false,
          error: "Google Apps Script returned a non-JSON response. Check deployment and permissions."
        })
      };
    }

    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store"
      },
      body: JSON.stringify(data)
    };
  } catch (error) {
    return {
      statusCode: 502,
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({
        success: false,
        error: "Cannot connect to Google Apps Script: " + (error.message || "network error")
      })
    };
  }
};
