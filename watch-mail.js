import fs from "fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SENDER = "zekielarroyo@gmail.com";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const TOKENS_FILE = path.join(__dirname, "tokens.json");


const { refresh_token } = JSON.parse(
	fs.readFileSync(TOKENS_FILE, "utf8")
);

async function refreshAccessToken() {
	const res = await fetch(
		"https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
		{
			method: "POST",
			headers: {
				"Content-Type":
				"application/x-www-form-urlencoded"
			},
			body: new URLSearchParams({
				client_id: "1e3f7ac9-62c7-403a-8aef-71fc6ba5b26a",
				client_secret: "tll8Q~kBIRW3ilRhzRCHEveX3z_zi0Epvfj~kcCl",
				grant_type: "refresh_token",
				refresh_token,
				// scope: "offline_access Mail.Read User.Read Mail.ReadBasic"
			})
		}
	);

	const token = await res.json();

	if (!res.ok) {
		throw new Error(
			`Token refresh failed: ${JSON.stringify(token, null, 2)}`
		);
	}

	if (token.access_token) {
		fs.writeFileSync(
			TOKENS_FILE,
			JSON.stringify(token, null, 2)
		);
	}

	return token.access_token;
}

async function main() {

    console.log("Refreshing token...");

    const accessToken = await refreshAccessToken();

    console.log("Token acquired:", !!accessToken);

    const res = await fetch(
        "https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false",
        {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        }
    );

    console.log("Graph status:", res.status);

    const data = await res.json();

    console.log("Messages found:", data.value?.length);

    const match = data.value?.find(
        m =>
            m.from?.emailAddress?.address?.toLowerCase() ===
            SENDER.toLowerCase()
    );

    console.log("Match found:", !!match);

	if (!match) {
		console.log("No email.");
		return;
	} else {
		const processed = JSON.parse(
			fs.readFileSync(path.join(__dirname, "processed.json"), "utf8")
		);

		if (processed.includes(match.id)) {
			return;
		}

		await fetch("http://127.0.0.1:18789/hooks/agent", {
			method: "POST",
			headers: {
				// "Authorization": `Bearer ${process.env.OPENCLAW_HOOK_TOKEN}`,
				"Authorization": `Bearer my-super-secret-token`,
				"Content-Type": "application/json"
			},
			body: JSON.stringify({
				message: `
A matching email was received.

Subject: ${match.subject}
Sender: ${match.from.emailAddress.address}
Message ID: ${match.id}

Read the email using the MS365 MCP and move attachment to desktop directory.
`
			})
		});

		processed.push(match.id);

		fs.writeFileSync(
			path.join(__dirname, "processed.json"),
			JSON.stringify(processed, null, 2)
		);
	}

	console.log("MATCH FOUND:", match.subject);

	// Trigger OpenClaw here
}
main();
