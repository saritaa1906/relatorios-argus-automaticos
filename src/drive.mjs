import fs from "node:fs";
import { google } from "googleapis";

export function createDrive(config) {
  const auth = new google.auth.OAuth2(config.googleClientId, config.googleClientSecret);
  auth.setCredentials({ refresh_token: config.googleRefreshToken });
  return google.drive({ version: "v3", auth });
}

export async function fileExists(drive, folderId, name) {
  const escaped = name.replaceAll("'", "\\'");
  const { data } = await drive.files.list({
    q: `'${folderId}' in parents and name='${escaped}' and trashed=false`,
    fields: "files(id,name)",
    spaces: "drive",
  });
  return (data.files || []).length > 0;
}

export async function uploadReport(drive, folderId, filePath, name) {
  const { data } = await drive.files.create({
    requestBody: { name, parents: [folderId] },
    media: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", body: fs.createReadStream(filePath) },
    fields: "id,name,webViewLink",
  });
  return data;
}

