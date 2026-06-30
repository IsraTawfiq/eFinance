const express = require("express");

const cors = require("cors");

const app = express();

app.use(cors());

app.use(express.json());

// ============ CONFIG (reads from Render Environment Variables) ============

const TENANT_ID     = process.env.TENANT_ID;

const CLIENT_ID     = process.env.CLIENT_ID;

const CLIENT_SECRET = process.env.CLIENT_SECRET;

const GROUP_ID      = process.env.WORKSPACE_ID;

const REPORT_ID     = process.env.REPORT_ID;

const PBI_USERNAME  = process.env.PBI_USERNAME;   // effective identity username

// Dataset ID (not in env, so hardcoded here)

const DATASET_ID    = "58186a37-ba4b-47fd-b7d3-e82c7d6118d3";

// Service Principal Object ID (NOT the Application ID!)

const SP_OBJECT_ID  = "9a13c5c0-b083-4668-ad18-35432d480d45";

// ==========================================================================

// Helper: get AAD access token (Service Principal / client credentials)

async function getAccessToken() {

  const res = await fetch(

    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,

    {

      method: "POST",

      headers: { "Content-Type": "application/x-www-form-urlencoded" },

      body: new URLSearchParams({

        grant_type:    "client_credentials",

        client_id:     CLIENT_ID,

        client_secret: CLIENT_SECRET,

        scope:         "https://analysis.windows.net/powerbi/api/.default"

      })

    }

  );

  const data = await res.json();

  if (!data.access_token) {

    throw new Error(data.error_description || "Failed to get access token");

  }

  return data.access_token;

}

// ---------- ENDPOINT 1: Get Gateway + Datasource IDs ----------

app.get("/get-ids", async (req, res) => {

  try {

    const token = await getAccessToken();

    const gwRes = await fetch("https://api.powerbi.com/v1.0/myorg/gateways", {

      headers: { Authorization: `Bearer ${token}` }

    });

    const gateways = await gwRes.json();

    const result = [];

    for (const gw of (gateways.value || [])) {

      const dsRes = await fetch(

        `https://api.powerbi.com/v1.0/myorg/gateways/${gw.id}/datasources`,

        { headers: { Authorization: `Bearer ${token}` } }

      );

      const ds = await dsRes.json();

      result.push({

        gatewayId:   gw.id,

        gatewayName: gw.name,

        datasources: (ds.value || []).map(d => ({

          datasourceId:   d.id,

          datasourceName: d.datasourceName,

          datasourceType: d.datasourceType

        }))

      });

    }

    res.json(result);

  } catch (e) {

    res.status(500).json({ error: e.message });

  }

});

// ---------- ENDPOINT 2: Grant ReadOverrideEffectiveIdentity ----------

app.get("/grant-override", async (req, res) => {

  try {

    const { gatewayId, datasourceId } = req.query;

    if (!gatewayId || !datasourceId) {

      return res.status(400).json({

        error: "Missing gatewayId or datasourceId in query. Run /get-ids first."

      });

    }

    const token = await getAccessToken();

    const grantRes = await fetch(

      `https://api.powerbi.com/v1.0/myorg/gateways/${gatewayId}/datasources/${datasourceId}/users`,

      {

        method: "POST",

        headers: {

          Authorization: `Bearer ${token}`,

          "Content-Type": "application/json"

        },

        body: JSON.stringify({

          identifier:            SP_OBJECT_ID,

          datasourceAccessRight: "ReadOverrideEffectiveIdentity",

          principalType:         "App"

        })

      }

    );

    if (grantRes.status === 200 || grantRes.status === 201) {

      return res.json({ success: true, message: "ReadOverrideEffectiveIdentity granted!" });

    }

    const errText = await grantRes.text();

    res.status(grantRes.status).json({

      success: false,

      status: grantRes.status,

      detail: errText

    });

  } catch (e) {

    res.status(500).json({ error: e.message });

  }

});

// ---------- ENDPOINT 3: Generate Embed Token (the main one) ----------

app.get("/embed", async (req, res) => {

  try {

    const token = await getAccessToken();

    const embedRes = await fetch(

      `https://api.powerbi.com/v1.0/myorg/groups/${GROUP_ID}/reports/${REPORT_ID}/GenerateToken`,

      {

        method: "POST",

        headers: {

          Authorization: `Bearer ${token}`,

          "Content-Type": "application/json"

        },

        body: JSON.stringify({

          accessLevel: "View",

          identities: [

            {

              username: PBI_USERNAME,

              datasets: [DATASET_ID]

            }

          ]

        })

      }

    );

    const embedData = await embedRes.json();

    if (!embedData.token) {

      return res.status(500).json({

        error: embedData.error?.message || embedData.message || "Failed to generate embed token",

        raw: embedData

      });

    }

    res.json({

      token:     embedData.token,

      embedUrl:  `https://app.powerbi.com/reportEmbed?reportId=${REPORT_ID}&groupId=${GROUP_ID}`,

      reportId:  REPORT_ID

    });

  } catch (e) {

    res.status(500).json({ error: e.message });

  }

});

// Health check

app.get("/", (req, res) => res.send("eFinance PBI backend is running."));

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
 
