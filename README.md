# cardBotTS

A bot that demonstrates [UAM](https://aka.ms/universal-actions-model) capabilities. But really this is a fun project 😍

This bot has been created using [Bot Framework](https://dev.botframework.com), it shows how to create a simple bot that accepts input from the user and echoes it back.

See it work!
![uam-image](./assets/uam.gif)

## Prerequisites

- [Node.js](https://nodejs.org) version 10.14.1 or higher

    ```bash
    # determine node version
    node --version
    ```
- [Microsoft 365 dev tenant](https://developer.microsoft.com/en-us/microsoft-365/dev-program?WT.mc_id=m365-35338-rwilliams)

- [A bot](https://dev.botframework.com/bots/) with ` Messaging endpoint` as the ngrok url appended with `api/messages`

## To run the bot

- Install modules

    ```bash
    npm install
    ```
- Start the bot

    ```bash
    npm start
    ```
- Start [ngrok](https://ngrok.com/) with command below

```bash
ngrok http-host-header=localhost:3978 
```
- Copy the ngrok url with https and paste it in the BOT configuration under `Messaging endpoint`

- Update the `botId`, `validDomains` and other necessary properties in the `manifest.json` file in the folder `appManifest`.

- Zip the three files (manifest.json and the icons) which is now your `Microsoft Teams app`

- Upload it to `Microsoft Teams` and use it in a `Team` or `Group chat` to further test it.


## Further reading

- [Bot Framework Documentation](https://docs.botframework.com?WT.mc_id=m365-35338-rwilliams)
- [Bot in Microsoft Teams](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/what-are-bots?WT.mc_id=m365-35338-rwilliams)
- [Universal action model](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/universal-actions-for-adaptive-cards/overview?WT.mc_id=m365-35338-rwilliams)
- [Bot generator](https://www.npmjs.com/package/generator-botbuilder)