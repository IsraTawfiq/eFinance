import express from "express";
import fetch from "node-fetch";
import cors from "cors";

const app = express();
app.use(cors());

const CONFIG = {
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  tenantId: process.env.TENANT_ID,
  workspaceId: process.env.WORKSPACE_ID,
  reportId: process.env.REPORT_ID,
};

async function getToken() {
  const res = await fetch(`https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    body: new URLSearchParams({
      grant_type: "client_credentials",
      client_id: CONFIG.clientId,
      client_secret: CONFIG.clientSecret,
      scope: "https://analysis.windows.net/powerbi/api/.default",
    }),
  });

  const data = await res.json();
  return data.access_token;
}

app.get("/embed", async (req, res) => {
  try {
    const token = await getToken();

    const r = await fetch(
      `https://api.powerbi.com/v1.0/myorg/groups/${CONFIG.workspaceId}/reports/${CONFIG.reportId}/GenerateToken`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ accessLevel: "view" }),
      }
    );

    const data = await r.json();

    res.json({
      embedToken: data.token,
      embedUrl: `https://app.powerbi.com/reportEmbed?reportId=${CONFIG.reportId}&groupId=${CONFIG.workspaceId}`,
      reportId: CONFIG.reportId,
    });

  } catch (e) {
    res.status(500).send("error");
  }
});

app.listen(10000, () => console.log("running"));
``
