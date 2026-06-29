export default async function handler(req, res) {
 res.setHeader("Access-Control-Allow-Origin", "*");
 
 const url = `https://login.microsoftonline.com/28a4cf11-f383-4845-bf32-a6cdadf71cd2/oauth2/v2.0/token`;
 
 const body = new URLSearchParams({
   grant_type:    "client_credentials",
   client_id:     "8915e96a-7e1a-42c9-85f4-b4b97234acf7",
   client_secret: process.env.PBI_SECRET,
   scope:         "https://analysis.windows.net/powerbi/api/.default"
 });

 const tokenRes = await fetch(url, {
   method: "POST",
   headers: { "Content-Type": "application/x-www-form-urlencoded" },
   body
 });
 const tokenData = await tokenRes.json();
 
 if (!tokenData.access_token) {
   return res.status(500).json({ error: tokenData.error_description });
 }

 const embedRes = await fetch(
   `https://api.powerbi.com/v1.0/myorg/groups/baf56118-c729-49bd-afd6-f473a0b7656c/reports/57c3e928-f0d4-4c8c-a441-9a4ac6c26a37/GenerateToken`,
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
   return res.status(500).json({ error: embedData.message });
 }

 res.status(200).json({ token: embedData.token });
}