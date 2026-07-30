import express from "express";
import open from "open";
import fs from "fs";

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID = process.env.TENANT_ID || "common";

const REDIRECT_URI = "http://localhost:3000/callback";

const app = express();

app.get("/callback", async (req, res) => {
    const code = req.query.code;

    const tokenRes = await fetch(
        `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
        {
            method: "POST",
            headers: {
                "Content-Type":
                    "application/x-www-form-urlencoded"
            },
            body: new URLSearchParams({
                client_id: CLIENT_ID,
                client_secret: CLIENT_SECRET,
                grant_type: "authorization_code",
                code,
                redirect_uri: REDIRECT_URI,
                // scope: "offline_access Mail.Read User.Read Mail.ReadBasic"
            })
        }
    );

    const tokens = await tokenRes.json();

    fs.writeFileSync(
        "./tokens.json",
        JSON.stringify(tokens, null, 2)
    );

    res.send("Success. You can close this window.");
    process.exit(0);
});

app.listen(3000);

const authUrl =
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
    new URLSearchParams({
        client_id: CLIENT_ID,
        response_type: "code",
        redirect_uri: REDIRECT_URI,
        response_mode: "query",
        scope: "offline_access Mail.Read User.Read"
    });

open(authUrl);
