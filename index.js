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

// ✅ get access token
async function getAccessToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: CONFIG.clientId,
        client_secret: CONFIG.clientSecret,
        scope: "https://analysis.windows.net/powerbi/api/.default",
      }),
    }
  );

  const data = await res.json();

  if (!res.ok) {
    throw new Error(JSON.stringify(data));
  }

  return data.access_token;
}

// ✅ MAIN ENDPOINT (فيه الحل)
app.get("/embed", async (req, res) => {
  try {
    const accessToken = await getAccessToken();

    // ✅ dataset ID (من error بتاعك)
    const datasetId = "58186a37-ba4b-47fd-b7d3-e82c7d6118d3";

    const response = await fetch(
      `https://api.powerbi.com/v1.0/myorg/groups/${CONFIG.workspaceId}/reports/${CONFIG.reportId}/GenerateToken`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          accessLevel: "view",

          // ✅ الحل هنا 👇
          identities: [
            {
              username: "user@demo.com",
              datasets: [datasetId]
            }
          ]
        }),
      }
    );

    const data = await response.json();

    if (!response.ok) {
      throw new Error(JSON.stringify(data));
    }

    res.json({
      embedToken: data.token,
      embedUrl: `https://app.powerbi.com/reportEmbed?reportId=${CONFIG.reportId}&groupId=${CONFIG.workspaceId}`,
      reportId: CONFIG.reportId,
    });

  } catch (e) {
    console.error(e);

    res.status(500).json({
      error: e.message
    });
  }
});

// ✅ test
app.get("/", (req, res) => {
  res.send("✅ API Running");
});

app.listen(10000, () => {
  console.log("✅ Server Running");
});
