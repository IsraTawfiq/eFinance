export default async function handler(req) {
 const TENANT_ID    = "28a4cf11-f383-4845-bf32-a6cdadf71cd2";
 const CLIENT_ID    = "8915e96a-7e1a-42c9-85f4-b4b97234acf7";
const CLIENT_SECRET = "_UD8Q-eJeFCMmnS6saMByn1TMEeZyvRT0hxXVc~I...";
 const GROUP_ID     = "baf56118-c729-49bd-afd6-f473a0b7656c";
 const REPORT_ID    = "57c3e928-f0d4-4c8c-a441-9a4ac6c26a37";
 const headers = {
   "Content-Type": "application/json",
   "Access-Control-Allow-Origin": "*"
 };
 try {
   // Step 1: Get Access Token
   const tokenRes = await fetch(
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
   const tokenData = await tokenRes.json();
   if (!tokenData.access_token) {
     return new Response(
       JSON.stringify({ error: tokenData.error_description || "Failed to get access token" }),
       { status: 500, headers }
     );
   }
   // Step 2: Get Embed Token
   const embedRes = await fetch(
     `https://api.powerbi.com/v1.0/myorg/groups/${GROUP_ID}/reports/${REPORT_ID}/GenerateToken`,
     {
       method: "POST",
       headers: {
         "Authorization": `Bearer ${tokenData.access_token}`,
         "Content-Type": "application/json"
       },
       body: JSON.stringify({ accessLevel: "View" })
     }
   );
   const embedData = await embedRes.json();
   if (!embedData.token) {
     return new Response(
       JSON.stringify({ error: embedData.message || "Failed to get embed token" }),
       { status: 500, headers }
     );
   }
   return new Response(
     JSON.stringify({ token: embedData.token }),
     { status: 200, headers }
   );
 } catch (err) {
   return new Response(
     JSON.stringify({ error: err.message }),
     { status: 500, headers }
   );
 }
}
