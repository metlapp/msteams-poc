# msteams-poc

Capstone project creating MS Teams POC app

# Installation

Below are the steps required to install, setup, and run the MS Teams App for a local development environment.

## Microsoft/Azure Account

1. Sign up for a Microsoft Developer account [here](https://developer.microsoft.com/en-us/microsoft-365/dev-program). This account will be used by Microsoft and Azure throughout the development process.

## VS Code Teams Toolkit

1. In VS Code, navigate to the Extensions Marketplace and look for "Teams Toolkit". You can also install it by navigating to the following link: [Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)
2. Once the extension has been installed, navigate to the newly created tab on the left side of VSCode. If you don't see the Teams Toolkit tab, you may need to restart VSCode.
3. In the Teams Toolkit tab, under the Accounts section, login to Microsoft 365 and Azure using your Microsoft developer account.

## MS Teams Bot

1. Clone the repo to the directory of your choice.
2. Open the newly installed repo folder in VSCode by going to File > Open Folder... then select the repo folder (msteams-poc). *NOTE* You may get prompted by VSCode to trust the developers of the files for that directory.
3. Once the folder has been loaded, you should be able to run the bot by pressing F5. This will open a new instance of Edge or Chrome and allow you to start testing the bot.
4. Once Edge or Chrome has opened, you will then be prompted on where to install the app. Click the dropdown beside the "Add" button and select "Add to a team". From there you can select what Team to install the app to.
5. Refer to the Known Issues/Best Practices section below for more information about running the bot and some possible issues.

# Known Issues/Best Practices

* The Teams Toolkit tends to sign you out of your accounts. This causes several issues including a subscription error. It's best practice if you check your accounts login status before attempting the run the bot.

* When multiple people are working on the bot at the same time, it's best to use your own Team for development. If multiple people add the bot to the same team it becomes difficult to differentiate between who's bot is who's. Even when you stop debugging and disable the bot, it stays in the team as an App (just without any functionality). One way to counteract this is to change the name of the bot in the .fx/manifest.source.json under name > short. I used Metl-tyson so that it was obviously my local debug version.

* When working with the bot, ensure the Metl API is running. Now that the API has been fully integrated with the bot it perform requests when events are fired in Teams like: Channels added/updated/deleted, Members added/removed, etc.

# Testing

You can use something like postman and send a POST request to the /proactivemessage endpoint with a body containing a type, text, targets, and choices (if testing a multi-choice question). For the targets field, you MUST send an entire object of either a channel or user, this can be obtained for the Organization in the Metl API. Refer to the samples below:

### User Target Sample

    {
        "id": "user-id",
        "name": "John Doe",
        "objectId": "object-id",
        "givenName": "John",
        "surname": "Doe",
        "email": "JohnDoe@email.onmicrosoft.com",
        "userPrincipalName": "JohnDoe@email.onmicrosoft.com",
        "tenantId": "tenant-id",
        "userRole": "user",
        "aadObjectId": "aadobject-id"
    }

### Channel Target Sample

    {
        "id": "channel-id",
        "name": "channel-name"
    }

## Simple Message

    {
        "type": "Message",
        "text": "This is a sample message directly from postman.",
        "targets": [
            {
                "id": "channel-id",
                "name": "channel-name"
            }
        ]
    }

## Yes/No

    {
        "type": "YesNo",
        "text": "Do you like dogs?",
        "targets": [
            {
                "id": "channel-id",
                "name": "channel-name"
            }
        ]
    }

## Multi-Choice

    {
        "type": "MultiChoice",
        "text": "What is your favourite household pet?",
        "choices": [
            {
                "title": "Dog",
                "value": "dog"
            },
            {
                "title": "Cat",
                "value": "cat"
            },
            {
                "title": "Fish",
                "value": "fish"
            }
        ],
        "targets": [
            {
                "id": "channel-id",
                "name": "channel-name"
            }
        ]
    }

## Text Block

    {
        "type": "TextBlock",
        "text": "Thoughts on dogs?",
        "targets": [
            {
                "id": "channel-id",
                "name": "channel-name"
            }
        ]
    }

## Number Only

    {
        "type": "Number",
        "text": "What is 3+2?",
        "targets": [
            {
                "id": "channel-id",
                "name": "channel-name"
            }
        ]
    }

# "Helpful" Resources

* [Proactive Messaging](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages?tabs=dotnet)
* [Card Types](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-reference)
* [Adaptive Cards](https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/what-are-cards)
* [Distribution](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-publish-overview)