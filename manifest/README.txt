Pim Teams App Package
=====================

To submit to IT for Teams admin approval, this folder needs:
  - manifest.json (done)
  - color.png  (192 x 192 px, full color Pim icon)
  - outline.png (32 x 32 px, white-on-transparent silhouette)

Once both PNGs are added, zip the THREE files (flat, not the folder) as pim-teams-app.zip:
  cd /Users/khunter2/Desktop/pim-bot/manifest
  zip ../pim-teams-app.zip manifest.json color.png outline.png

Then give pim-teams-app.zip to IT so they can upload it in Teams Admin Center →
Manage apps → Upload new app → Submit for approval.

IMPORTANT: The bot ID in manifest.json (331d7882-e0db-41d7-af70-92571209a943)
must match the Azure App Registration that IT created and the Azure Bot resource
(pim-athome-creative-2026).
