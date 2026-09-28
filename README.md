# Write Buddy for Outlook

Write Buddy is an Outlook add-in that lets users improve, rewrite, fact-check, and translate email text without leaving Outlook.

The project was built to bring an AI writing workflow directly into a tool people already use every day instead of forcing them to copy and paste text into a separate chat application.

## Features

- Rewrite selected text professionally
- Improve clarity while preserving the original meaning
- Fact-check and revise selected content
- Make writing more persuasive
- Translate into 24+ languages
- Review the generated version before inserting it
- Replace selected email text directly inside Outlook
- Enable or disable individual tools
- Store configuration through Outlook's add-in settings

## How it works

1. The user opens the Write Buddy task pane in Outlook
2. Selected email text is read through the Office JavaScript API
3. The user chooses a writing action
4. The add-in sends the request to the configured AI service
5. The result is shown for review
6. The user can insert the revised text back into the email

## Stack

- Microsoft Outlook Add-in
- Office JavaScript API
- JavaScript
- HTML / CSS
- OpenAI API
- Office add-in manifest
- Local HTTPS development tooling

## Project structure

```
write-buddy-outlook/
├── manifest.xml
├── taskpane.html
├── taskpane.js
├── taskpane.css
├── package.json
└── README.md
```

## Local development

Outlook add-ins require HTTPS during development.

```bash
npm install
npm start
```

Then sideload `manifest.xml` into Outlook and open the task pane.

Configuration such as API credentials should be supplied at runtime rather than committed to source control.

## Why this project matters

Write Buddy is an example of embedding AI into an existing employee workflow instead of building another standalone AI interface. The goal is to reduce friction: select text, choose an action, review it, and continue working.

---

Built by [Joel Haymour](https://github.com/joelhaymour).
