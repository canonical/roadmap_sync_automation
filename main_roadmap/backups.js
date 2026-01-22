const RUN_BACKUP_FOLDER_ID = "1B37pAPfBXAsTlSD3azt4FY-mrNhaa3jT";
const RUN_RETENTION_DAYS = 90;

function cleanupBackups() {
  const start = Date.now();

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseNamePrefix = `${ss.getName()} - Backup `;

  const cutoff = new Date(Date.now() - RUN_RETENTION_DAYS * 24 * 60 * 60 * 1000);
  const folder = DriveApp.getFolderById(RUN_BACKUP_FOLDER_ID);
  const files = folder.getFiles();

  let checked = 0;
  let deleted = 0;

  while (files.hasNext()) {
    const f = files.next();
    checked++;

    const name = f.getName();
    if (!name.startsWith(baseNamePrefix)) continue;

    if (f.getDateCreated() < cutoff) {
      f.setTrashed(true);
      deleted++;
    }
  }

  Logger.log(`cleanupRunBackups_: checked=${checked}, deleted=${deleted}, time=${Date.now() - start} ms`);
}
