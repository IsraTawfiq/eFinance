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
  datasetId: "58186a37-ba4b-47fd-b7d3-e82c7d6118d3"
};

async function getAccessToken() {
  const response = await fetch(
    `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: CONFIG.clientId,
        client_secret: CONFIG.clientSecret,
        scope: "https://analysis.windows.net/powerbi/api/.default"
      })
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(JSON.stringify(data));
  }

  return data.access_token;
}

app.get("/", (req, res) => {
  res.send("✅ API Running");
});

app.get("/embed", async (req, res) => {
  try {
    const accessToken = await getAccessToken();

    const response = await fetch(
      `https://api.powerbi.com/v1.0/myorg/groups/${CONFIG.workspaceId}/reports/${CONFIG.reportId}/GenerateToken`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          accessLevel: "view",
          identities: [
            {
              username: "powerbi_user@efinance.com.eg",
              datasets: [CONFIG.datasetId]
            }
          ]
        })
      }
    );

    const data = await response.json();

    console.log(JSON.stringify(data, null, 2));

    if (!response.ok) {
      return res.status(500).json(data);
    }

    return res.json({
      reportId: CONFIG.reportId,
      embedUrl: `https://app.powerbi.com/reportEmbed?reportId=${CONFIG.reportId}&groupId=${CONFIG.workspaceId}`,
      embedToken: data.token
    });

  } catch (error) {
    return res.status(500).json({
      error: error.message
    });
  }
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
  console.log("✅ Server Running");
});
