# Introduction
This script is a webhook for google workspace group. Which extracts the latest notices from `https://vulms.vu.edu.pk/NoticeBoard/NoticeBoard2.aspx` link and post them in the google chat.
Contribution will be appreciated to make it more flexible.

## Process to add it in your Google workspace: -
1. Click on your google workspace group and go to `Apps & intergraions` section.
2. Then click `Add webhooks`, given under webhook section.
3. Give it any preferable name. Adding `Avatar URL` is optional.
4. After this, click on `save`.
5. Then click `copy link` shown after you click three dots next to your added webhook.
6. Go to `https://script.google.com/`.
7. Click `New Project` on top left corner.
8. Give your project any name.
9. Copy and paste the given script in the editor there.
10. Go to `Project settings` from the left menu bar.
11. At the bottom. add script properties with the property name: `GOOGLE_CHAT_WEBHOOK_URL` and replace value with the copied link from the step 5.
12. Switch to editor and press `Ctrl+C` to save your project.

*Here you've your notice board posts with perfect format in your google chats!*
