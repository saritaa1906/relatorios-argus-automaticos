const required = ["ARGUS_TOKEN", "GOOGLE_CLIENT_ID", "GOOGLE_CLIENT_SECRET", "GOOGLE_REFRESH_TOKEN", "GOOGLE_DRIVE_FOLDER_ID"];

export function loadConfig() {
  const dryRun = String(process.env.DRY_RUN || "false").toLowerCase() === "true";
  const missing = required.filter((name) => !process.env[name] && !(dryRun && name.startsWith("GOOGLE_")));
  if (missing.length) throw new Error(`Secrets ausentes: ${missing.join(", ")}`);
  return {
    argusToken: process.env.ARGUS_TOKEN,
    googleClientId: process.env.GOOGLE_CLIENT_ID,
    googleClientSecret: process.env.GOOGLE_CLIENT_SECRET,
    googleRefreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    driveFolderId: process.env.GOOGLE_DRIVE_FOLDER_ID,
    trackingDate: process.env.START_TRACKING_DATE || "2026-08-01",
    dryRun,
  };
}

