// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { ActivityTypes, CardFactory, ActionTypes } = require('botbuilder');
const { DialogSet, WaterfallDialog, DateTimePrompt, DialogTurnStatus } = require('botbuilder-dialogs');
const { LuisRecognizer } = require('botbuilder-ai');
const allHolidays = JSON.parse(require("./JSON_Files/Holidays.json"));
//#region  state properties accessors
const DIALOG_STATE_PROPERTY_ACCESSOR = 'dialogStatePropertyAccessor';
const USER_PROFILE_PROPERTY_ACCESSOR = 'userProfilePropertyAccessor';
const CONVERSATION_DATA_PROPERTY_ACCESSOR = 'conversationDataPropertyAccessor';
//#endregion
const MAX_FLEXI_LEAVE_COUNT = 3;
const MAX_LEAVE_COUNT = 27;
const CURRENT_YEAR = new Date().getFullYear();
const FIRST_DAY_OF_YEAR = new Date(CURRENT_YEAR, 0, 1);
const LAST_DAY_OF_YEAR = new Date(CURRENT_YEAR, 11, 31);
const CURRENT_DATE = new Date();
const FLEXI_HOLIDAY_MESSAGE_REGEX = new RegExp(`^{{{(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)}}} - {{{${CURRENT_YEAR}-(0|1)\\d-(0|1|2|3)\\d}}} - {{{.+}}}$`);
//#region Cards
const holidayAdaptiveCard = require('./cards/HolidayAdaptiveCard.json');
const holidayAdaptiveCardTwo = require('./cards/HolidayAdaptiveCardTwo.json');
const holidayAdaptiveCardColumnSet = require('./cards/HolidayAdaptiveCardColumnSet.json');
const holidayAdaptiveCardTwoColumnSet = require('./cards/HolidayAdaptiveCardTwoColumnSet')
const helpCard = require('./cards/HelpCard.json')
//#endregion
//#region Dialogs IDs
const LEAVE_DATE_DIALOG = 'LeaveDateDialog'
const START_DATE_PROMPT = 'startDatePrompt';
const END_DATE_PROMPT = 'endDatePrompt';
//#endregion
//#region supported LUIS Intents
const CANCEL_INTENT = 'Cancel';
const GREETING_INTENT = 'Greeting';
const HELP_INTENT = 'Help';
const LEAVE_REQUEST_INTENT = 'LeaveRequest';
const NONE_INTENT = 'None';
const SHOW_ALL_FLEXIBLE_HOLIDAYS_INTENT = 'ShowAllFlexiHolidays';
const SHOW_ALL_HOLIDAYS_INTENT = 'ShowAllHolidays';
const SHOW_MY_LEAVES_INTENT = 'ShowMyLeaves';
const SHOW_MY_FLEXI_LEAVES_INTENT = "ShowMyFlexiLeaves"
//#endregion
class LuisBot {

    constructor(application, luisPredictionOptions, conversationState, userState) {
        this.conversationData = conversationState.createProperty(CONVERSATION_DATA_PROPERTY_ACCESSOR);
        this.userProfile = userState.createProperty(USER_PROFILE_PROPERTY_ACCESSOR);
        this.conversationState = conversationState;
        this.userState = userState;

        this.dialogState = this.conversationState.createProperty(DIALOG_STATE_PROPERTY_ACCESSOR);
        this.dialogSet = new DialogSet(this.dialogState);

        this.dialogSet.add(new DateTimePrompt(START_DATE_PROMPT, this.dateValidator));
        this.dialogSet.add(new DateTimePrompt(END_DATE_PROMPT, this.dateValidator));
        // Define the steps of the waterfall dialog and add it to the set.
        this.dialogSet.add(
            new WaterfallDialog(LEAVE_DATE_DIALOG, [
                this.promptForStartDate.bind(this),
                this.promptForEndDate.bind(this),
                this.acknowledgeLeave.bind(this)
            ])
        );
        this.luisRecognizer = new LuisRecognizer(application, luisPredictionOptions, true);
    }
    async promptForStartDate(stepContext) {
        return await stepContext.prompt(START_DATE_PROMPT, {
            prompt: 'What is start date of Leave?',
            retryPrompt: 'Please retry....What is start duration of Leave?'
        });
    }
    async promptForEndDate(stepContext) {
        stepContext.values.startDate = stepContext.result;
        return await stepContext.prompt(END_DATE_PROMPT, {
            prompt: 'What is end date of Leave?. Enter same date as start date if you want leave for one day only',
            retryPrompt: 'Please retry....What is end date of Leave? Enter same date as start date if you want leave for one day only'
        });
    }
    async dateValidator(promptContext) {
        // Check whether the input could be recognized as an integer.
        if (!promptContext.recognized.succeeded) {
            await promptContext.context.sendActivity(
                "I'm sorry, I do not understand. Please enter the date for your leave."
            );
            return false;
        }
        // Check whether any of the recognized date-times are appropriate,
        // and if so, return the first appropriate date-time.
        let value = null;
        promptContext.recognized.value.forEach(candidate => {
            const date = new Date(candidate.value || candidate.start);
            if (FIRST_DAY_OF_YEAR <= date && date <= LAST_DAY_OF_YEAR) {
                value = candidate;
            }
        });
        if (value) {
            promptContext.recognized.value = [value];
            return true;
        }

        await promptContext.context.sendActivity(
            "I'm sorry, date should lie in current year"
        );
        return false;
    }
    async acknowledgeLeave(stepContext) {
        // Retrieve the reservation date.
        stepContext.values.endDate = stepContext.result;
        // Return the collected information to the parent context.
        return await stepContext.endDialog({
            date: stepContext.values,
        });
    }

    async onTurn(turnContext) {
        if (turnContext.activity.type === ActivityTypes.Message) {
            let userProfile = await this.userProfile.get(turnContext, {});
            let conversationData = await this.conversationData.get(turnContext, { promptedForUserName: false });


            const dialogContext = await this.dialogSet.createContext(turnContext);
            const dialogContextResults = await dialogContext.continueDialog();
            if (dialogContextResults.status != DialogTurnStatus.empty && dialogContextResults.status != DialogTurnStatus.cancelled) {
                switch (dialogContextResults.status) {

                    case DialogTurnStatus.waiting:
                        // If there is an active dialog, we don't need to do anything here.
                        break;
                    case DialogTurnStatus.complete:
                        // If we just finished the dialog, capture and display the results.
                        const leaveInfo = dialogContextResults.result.date;
                        const startDate = new Date(leaveInfo.startDate[0].value);
                        const endDate = new Date(leaveInfo.endDate[0].value);
                        let leaveAddStatus = this.addUserLeaves(startDate, endDate, userProfile);
                        if (leaveAddStatus.totalLeaveAddedToBalance > 0) {
                            await turnContext.sendActivity(`Leave successfully opted. Your remaining leave balance is ${MAX_LEAVE_COUNT - userProfile.holidays.length}`);
                        }
                        else if (leaveAddStatus.message !== '') {
                            await turnContext.sendActivity(leaveAddStatus.message);
                        }
                        else {
                            await turnContext.sendActivity(`Leaves could not be added. Please try again`);
                        }
                        break;
                }
                await this.userProfile.set(turnContext, userProfile);
                await this.userState.saveChanges(turnContext);
                await this.conversationData.set(turnContext, conversationData);
                await this.conversationState.saveChanges(turnContext);
                return;
            }


            if (!userProfile.name) {
                if (conversationData.promptedForUserName) {
                    userProfile.name = turnContext.activity.text;
                    // or start main dialog
                    await turnContext.sendActivity(
                        `Hi ${userProfile.name}. This bot helps you to view and manage your planned and flexible leaves. Type 'Help' anytime to get a quick glimpse of sample commands and information you can ask for...`
                    );
                    conversationData.promptedForUserName = false;
                } else {
                    await turnContext.sendActivity('Hi User...What is your name?');
                    conversationData.promptedForUserName = true;
                }
                await this.userProfile.set(turnContext, userProfile);
                await this.userState.saveChanges(turnContext);
                await this.conversationData.set(turnContext, conversationData);
                await this.conversationState.saveChanges(turnContext);
            }
            else if (conversationData.flexibleHolidayListDisplayed && FLEXI_HOLIDAY_MESSAGE_REGEX.test(turnContext.activity.text)) {
                var tokens = turnContext.activity.text.split(' - ');
                var day = tokens[0].substr(3).substr(0, tokens[0].length - 6);
                var date = tokens[1].substr(3).substr(0, tokens[1].length - 6);
                var name = tokens[2].substr(3).substr(0, tokens[2].length - 6);
                if (userProfile.flexiHolidays) {
                    if (userProfile.flexiHolidays.length == 3) {
                        await turnContext.sendActivity(`Sorry ${userProfile.name}.. You have already opted your ${MAX_FLEXI_LEAVE_COUNT} flexi leaves`);
                    }
                    else {
                        let alreadyExist = false;
                        for (let index = 0; index < userProfile.flexiHolidays.length; index++) {
                            const userFlexiHoliday = userProfile.flexiHolidays[index];
                            if (userFlexiHoliday.date == date) {
                                alreadyExist = true;
                                break;
                            }
                        }
                        if (alreadyExist) {
                            await turnContext.sendActivity(`You can not opt for same day twice. Please choose a different day`);
                        }
                        else {
                            userProfile.flexiHolidays.push({ day: day, date: date, name: name })
                            await turnContext.sendActivity(`Flexible Leave successfully opted. Your remaining flexi leave balance is ${MAX_FLEXI_LEAVE_COUNT - userProfile.flexiHolidays.length}`);
                        }
                    }
                }
                else {
                    userProfile.flexiHolidays = [{ day: day, date: date, name: name }];
                    await turnContext.sendActivity(`Flexible Leave successfully opted. Your remaining flexi leave balance is ${MAX_FLEXI_LEAVE_COUNT - userProfile.flexiHolidays.length}`);
                }
                this.conversationData.flexibleHolidayListDisplayed = false;
                await this.userProfile.set(turnContext, userProfile);
                await this.userState.saveChanges(turnContext);
                await this.conversationData.set(turnContext, conversationData);
                await this.conversationState.saveChanges(turnContext);
            }
            else {
                //since Luis Recognizer takes some time so send "typing" activity first 
                await turnContext.sendActivity({ type: 'typing' });

                const results = await this.luisRecognizer.recognize(turnContext);
                const topIntent = results.luisResult.topScoringIntent;
                let parsedDate = this.extractDateFromQuery(results);
                if (topIntent.intent == HELP_INTENT) {
                    const reply = {
                        text: 'Help',
                        attachments: [CardFactory.adaptiveCard(helpCard)]
                    };
                    await turnContext.sendActivity(reply);
                }
                else if (topIntent.intent == CANCEL_INTENT) {
                    await turnContext.sendActivity(`You cancelled the current operation. Type help for more information`);
                }
                else if (topIntent.intent == GREETING_INTENT) {
                    await turnContext.sendActivity(`Hi ${userProfile.name}. How may I help you ? Type Help for a quick help`);
                }
                else if (topIntent.intent == LEAVE_REQUEST_INTENT) {
                    if (parsedDate.isDateComponentPresent) {
                        if (parsedDate.isDateComponentValid) {
                            let leaveAddStatus = this.addUserLeaves(parsedDate.startDate, parsedDate.endDate, userProfile);
                            if (leaveAddStatus.totalLeaveAddedToBalance > 0) {
                                await turnContext.sendActivity(`Leave successfully opted. Your remaining leave balance is ${MAX_LEAVE_COUNT - userProfile.holidays.length}`);
                            }
                            else if (leaveAddStatus.message !== '') {
                                await turnContext.sendActivity(leaveAddStatus.message);
                            }
                            else {
                                await turnContext.sendActivity(`Leaves could not be added. Please try again`);
                            }
                            await this.userProfile.set(turnContext, userProfile);
                            await this.userState.saveChanges(turnContext);
                        }
                        else {
                            await turnContext.sendActivity('Sorry I could not interpret that date... Please Try again');
                        }
                    }
                    else {
                        if (dialogContextResults.status == DialogTurnStatus.empty || dialogContextResults.status == DialogTurnStatus.cancelled) {
                            await dialogContext.beginDialog(LEAVE_DATE_DIALOG);
                            await this.userProfile.set(turnContext, userProfile);
                            await this.userState.saveChanges(turnContext);
                            await this.conversationData.set(turnContext, conversationData);
                            await this.conversationState.saveChanges(turnContext);
                        }
                    }

                }
                else if (topIntent.intent == SHOW_ALL_FLEXIBLE_HOLIDAYS_INTENT) {
                    if (parsedDate.isDateComponentPresent) {
                        if (parsedDate.isDateComponentValid) {
                            parsedDate.endDate.setDate(parsedDate.endDate.getDate() - 1);
                            const selectedFlexibleHolidayListButtons = this.getFlexibleHolidayListButtons(parsedDate.startDate, parsedDate.endDate, userProfile);
                            if (selectedFlexibleHolidayListButtons.length > 0) {
                                const heroCardInstance = CardFactory.heroCard("Available Flexible Leaves", undefined, selectedFlexibleHolidayListButtons, { text: 'Click on a button to opt it as a flexi leave' });
                                const reply = { type: ActivityTypes.Message, attachments: [heroCardInstance] };
                                await turnContext.sendActivity(reply);
                                conversationData.flexibleHolidayListDisplayed = true;
                                await this.conversationData.set(turnContext, conversationData);
                                await this.conversationState.saveChanges(turnContext);
                            }
                            else {
                                await turnContext.sendActivity(`No holidays were found corresponding to your query`);
                            }
                        }
                        else {
                            await turnContext.sendActivity('Sorry I could not interpret that date... Please Try again');
                        }
                    }
                    else {
                        const allFlexibleHolidayListButtons = this.getFlexibleHolidayListButtons(CURRENT_DATE, LAST_DAY_OF_YEAR, userProfile);
                        if (allFlexibleHolidayListButtons.length > 0) {
                            const heroCardInstance = CardFactory.heroCard("Available Flexible Leaves", undefined, allFlexibleHolidayListButtons, { text: 'Click on a button to opt it as a flexi leave' });
                            const reply = { type: ActivityTypes.Message, attachments: [heroCardInstance] };
                            await turnContext.sendActivity(reply);
                            conversationData.flexibleHolidayListDisplayed = true;
                            await this.conversationData.set(turnContext, conversationData);
                            await this.conversationState.saveChanges(turnContext);
                        }
                        else {
                            await turnContext.sendActivity(`No holidays were found corresponding to your query`);
                        }

                    }
                }
                else if (topIntent.intent == SHOW_ALL_HOLIDAYS_INTENT) {
                    if (parsedDate.isDateComponentPresent) {
                        if (parsedDate.isDateComponentValid) {
                            parsedDate.endDate.setDate(parsedDate.endDate.getDate() - 1);
                            const selectedHolidayAdaptiveCard = this.getCustomHolidayAdaptiveCard(parsedDate.startDate, parsedDate.endDate);
                            if (selectedHolidayAdaptiveCard.body.length > 2) {
                                const reply = {
                                    attachments: [CardFactory.adaptiveCard(selectedHolidayAdaptiveCard)]
                                };
                                await turnContext.sendActivity(reply);
                            }
                            else {
                                await turnContext.sendActivity(`No holidays were found corresponding to your query`);
                            }
                        }
                        else {
                            await turnContext.sendActivity('Sorry I could not interpret that date... Please Try again');
                        }
                    }
                    else {
                        //send list of all holidays
                        const allHolidayAdaptiveCard = this.getCustomHolidayAdaptiveCard(CURRENT_DATE, LAST_DAY_OF_YEAR);
                        if (allHolidayAdaptiveCard.body.length > 2) {
                            const reply = {
                                attachments: [CardFactory.adaptiveCard(allHolidayAdaptiveCard)]
                            };
                            await turnContext.sendActivity(reply);
                        }
                        else {
                            await turnContext.sendActivity(`No holidays were found corresponding to your query`);
                        }

                    }
                }
                else if (topIntent.intent == SHOW_MY_LEAVES_INTENT) {
                    if (parsedDate.isDateComponentPresent) {
                        if (parsedDate.isDateComponentValid) {
                            const selectedUserHolidayAdaptiveCard = this.getUserHolidayAdaptiveCard(parsedDate.startDate, parsedDate.endDate, userProfile);
                            if (selectedUserHolidayAdaptiveCard.body.length > 2) {
                                const reply = {
                                    attachments: [CardFactory.adaptiveCard(selectedUserHolidayAdaptiveCard)]
                                };
                                await turnContext.sendActivity(reply);
                            }
                            else {
                                await turnContext.sendActivity(`No holidays were found corresponding to your query`);
                            }
                        }
                        else {
                            await turnContext.sendActivity('Sorry I could not interpret that date... Please Try again');
                        }
                    }
                    else {
                        //send list of all opted holidays
                        const allUserHolidayAdaptiveCard = this.getUserHolidayAdaptiveCard(CURRENT_DATE, LAST_DAY_OF_YEAR, userProfile);
                        if (allUserHolidayAdaptiveCard.body.length > 2) {
                            const reply = {
                                attachments: [CardFactory.adaptiveCard(allUserHolidayAdaptiveCard)]
                            };
                            await turnContext.sendActivity(reply);
                        }
                        else {
                            await turnContext.sendActivity(`No holidays were found corresponding to your query`);
                        }

                    }
                }
                else if (topIntent.intent == SHOW_MY_FLEXI_LEAVES_INTENT) {
                    if (parsedDate.isDateComponentPresent) {
                        if (parsedDate.isDateComponentValid) {
                            const selectedFlexiHolidayAdaptiveCard = this.getFlexiHolidayAdaptiveCard(parsedDate.startDate, parsedDate.endDate, userProfile);
                            if (selectedFlexiHolidayAdaptiveCard.body.length > 2) {
                                const reply = {
                                    attachments: [CardFactory.adaptiveCard(selectedFlexiHolidayAdaptiveCard)]
                                };
                                await turnContext.sendActivity(reply);
                            }
                            else {
                                await turnContext.sendActivity(`No holidays were found corresponding to your query`);
                            }
                        }
                        else {
                            await turnContext.sendActivity('Sorry I could not interpret that date... Please Try again');
                        }
                    }
                    else {
                        //send list of all flexi holidays
                        const allFlexiHolidayAdaptiveCard = this.getFlexiHolidayAdaptiveCard(CURRENT_DATE, LAST_DAY_OF_YEAR, userProfile);
                        if (allFlexiHolidayAdaptiveCard.body.length > 2) {
                            const reply = {
                                attachments: [CardFactory.adaptiveCard(allFlexiHolidayAdaptiveCard)]
                            };
                            await turnContext.sendActivity(reply);
                        }
                        else {
                            await turnContext.sendActivity(`No holidays were found corresponding to your query`);
                        }

                    }
                }
                else if (topIntent.intent === NONE_INTENT) {
                    // If the top scoring intent was "None" tell the user no valid intents were found and provide help.
                    await turnContext.sendActivity(`I am sorry. I could not get you input. Please try again... Type 'Help' anytime to get a quick glimpse of sample commands and information you can ask for...`);
                }
            }
        }
        else if (turnContext.activity.type === ActivityTypes.ConversationUpdate &&
            turnContext.activity.recipient.id !== turnContext.activity.membersAdded[0].id) {
            const conversationData = await this.conversationData.get(turnContext, { promptedForUserName: false });
            await turnContext.sendActivity('Welcome User. I am Leave Management Bot. May I know your name?');
            conversationData.promptedForUserName = true;
            await this.conversationData.set(turnContext, conversationData);
            await this.conversationState.saveChanges(turnContext);
        }
        else if (turnContext.activity.type !== ActivityTypes.ConversationUpdate) {
            // Respond to all other Activity types.
            await turnContext.sendActivity("Unknown Input... Please try again");
            console.log(`[${turnContext.activity.type}]-type activity detected.`);
        }
    }
    addUserLeaves(startDate, endDate, userProfile) {
        let currDate = startDate;
        let leaveAddStatus = { totalLeaveAddedToBalance: 0, message: '' };
        do {
            let isPublicHoliday = false;
            for (let index = 0; index < allHolidays.length; index++) {
                let publicHoliday = allHolidays[index];
                let publicHolidayDate = new Date(publicHoliday.Date);
                if (currDate.toLocaleDateString() == publicHolidayDate.toLocaleDateString()) {
                    isPublicHoliday = true;
                    leaveAddStatus.message += `Your opted leave for ${publicHolidayDate.toLocaleDateString()} can not be added because it is a public holiday\n`
                    break;
                }
            }
            let isFlexiHoliday = false;
            if (userProfile.flexiHolidays) {
                for (let index = 0; index < userProfile.flexiHolidays.length; index++) {
                    let flexiHoliday = userProfile.flexiHolidays[index];
                    let flexiHolidayDate = new Date(flexiHoliday.date);
                    if (currDate.toLocaleDateString() == flexiHolidayDate.toLocaleDateString()) {
                        isFlexiHoliday = true;
                        leaveAddStatus.message += `Your opted leave for ${flexiHolidayDate.toLocaleDateString()} can not be added because it is already covered under flexible holidays\n`
                        break;
                    }
                }
            }
            var dayOfWeek = currDate.getDay();
            if(dayOfWeek == 0 || dayOfWeek == 6)
            {
                leaveAddStatus.message += `Your opted leave for ${currDate.toLocaleDateString()} can not be added because it is on weekend\n`
            }
            if (!isPublicHoliday && !isFlexiHoliday && dayOfWeek != 0 && dayOfWeek != 6) {
                if (userProfile.holidays) {
                    if (userProfile.holidays.length == MAX_LEAVE_COUNT) {
                        leaveAddStatus.message += `Sorry ${userProfile.name}.. You have exceeded your ${MAX_LEAVE_COUNT} leaves quota\n`
                        break;
                    }
                    else {
                        let alreadyExist = false;
                        for (let index = 0; index < userProfile.holidays.length; index++) {
                            const userHoliday = userProfile.holidays[index];
                            if (userHoliday.date == currDate.toLocaleDateString()) {
                                alreadyExist = true;
                                if (leaveAddStatus)
                                    leaveAddStatus.message += `You have already opted leave for ${userHoliday.date}\n`
                                break;
                            }
                        }
                        if (!alreadyExist) {
                            userProfile.holidays.push({ date: currDate.toLocaleDateString() });
                            leaveAddStatus.totalLeaveAddedToBalance++;
                        }
                    }
                }
                else {
                    userProfile.holidays = [{ date: currDate.toLocaleDateString() }];
                    leaveAddStatus.totalLeaveAddedToBalance = 1;
                }
            }
            currDate.setDate(currDate.getDate() + 1);
        }
        while (currDate < endDate)
        return leaveAddStatus;
    }
    getCustomHolidayAdaptiveCard(startDate, endDate) {
        let customHolidayAdaptiveCard = JSON.parse(JSON.stringify(holidayAdaptiveCard));
        for (let index = 0; index < allHolidays.length; index++) {
            const holiday = allHolidays[index];
            var holidayDate = new Date(holiday.Date);
            if (holidayDate >= startDate && holidayDate <= endDate) {
                let columnSet = JSON.parse(JSON.stringify(holidayAdaptiveCardColumnSet));
                columnSet.columns[0].items[0].text = holiday.Day;
                columnSet.columns[1].items[0].text = holiday.Date;
                columnSet.columns[2].items[0].text = holiday.Name;
                columnSet.columns[3].items[0].text = holiday.IsFlexible === "True" ? "Flexible" : "Fixed";
                customHolidayAdaptiveCard.body.push(columnSet);
            }
        }
        return customHolidayAdaptiveCard;
    }
    getFlexiHolidayAdaptiveCard(startDate, endDate, userProfile) {
        let customHolidayAdaptiveCard = JSON.parse(JSON.stringify(holidayAdaptiveCard));
        if (userProfile.flexiHolidays) {
            for (let index = 0; index < userProfile.flexiHolidays.length; index++) {
                const holiday = userProfile.flexiHolidays[index];
                var holidayDate = new Date(holiday.date);
                if (holidayDate >= startDate && holidayDate <= endDate) {
                    let columnSet = JSON.parse(JSON.stringify(holidayAdaptiveCardColumnSet));
                    columnSet.columns[0].items[0].text = holiday.day;
                    columnSet.columns[1].items[0].text = holiday.date;
                    columnSet.columns[2].items[0].text = holiday.name;
                    columnSet.columns[3].items[0].text = "Flexible";
                    customHolidayAdaptiveCard.body.push(columnSet);
                }
            }
        }
        return customHolidayAdaptiveCard;
    }
    getUserHolidayAdaptiveCard(startDate, endDate, userProfile) {
        let customHolidayAdaptiveCard = JSON.parse(JSON.stringify(holidayAdaptiveCardTwo));
        if (userProfile.holidays) {
            const weekDays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
            for (let index = 0; index < userProfile.holidays.length; index++) {
                const holiday = userProfile.holidays[index];
                var holidayDate = new Date(holiday.date);
                var dayOfWeek = holidayDate.getDay();
                if (holidayDate >= startDate && holidayDate <= endDate) {
                    let columnSet = JSON.parse(JSON.stringify(holidayAdaptiveCardTwoColumnSet));
                    columnSet.columns[0].items[0].text = weekDays[dayOfWeek];
                    columnSet.columns[1].items[0].text = holiday.date;
                    customHolidayAdaptiveCard.body.push(columnSet);
                }
            }
        }
        return customHolidayAdaptiveCard;
    }
    getFlexibleHolidayListButtons(startDate, endDate, userProfile) {
        let buttons = [];

        for (let index = 0; index < allHolidays.length; index++) {
            const holiday = allHolidays[index];
            var holidayDate = new Date(holiday.Date);
            if (holidayDate >= startDate && holidayDate <= endDate && holiday.IsFlexible === "True") {
                if (userProfile.flexiHolidays) {
                    for (let index = 0; index < userProfile.flexiHolidays.length; index++) {
                        const flexiHoliday = userProfile.flexiHolidays[index];
                        if (flexiHoliday.date !== holiday.Date) {
                            buttons.push({
                                type: ActionTypes.PostBack,
                                title: `[${holiday.Day}] - [${holiday.Date}] - [${holiday.Name}]`,
                                value: `{{{${holiday.Day}}}} - {{{${holiday.Date}}}} - {{{${holiday.Name}}}}`
                            });
                            break;
                        }
                    }
                }
                else {
                    buttons.push({
                        type: ActionTypes.PostBack,
                        title: `[${holiday.Day}] - [${holiday.Date}] - [${holiday.Name}]`,
                        value: `{{{${holiday.Day}}}} - {{{${holiday.Date}}}} - {{{${holiday.Name}}}}`
                    });
                }

            }
        }
        return buttons;
    }
    
    isObjectEmpty(obj) {
        return Object.getOwnPropertyNames(obj).length == 0;
    }

    extractDateFromQuery(results) {
        let parsedDate = {};
        if (!this.isObjectEmpty(results.entities[`$instance`])) {
            parsedDate.isDateComponentPresent = true;
            var dateObjects = results.luisResult.entities[0].resolution.values;
            dateObjects.forEach(dateObj => {
                if (dateObj.type === 'daterange') {
                    let startDate = new Date(dateObj.start);
                    let endDate = new Date(dateObj.end);
                    if (startDate >= FIRST_DAY_OF_YEAR && endDate <= LAST_DAY_OF_YEAR) {
                        parsedDate.isDateComponentValid = true;
                        parsedDate.startDate = startDate;
                        parsedDate.endDate = endDate;
                    }
                }
                else if (dateObj.type === 'date') {
                    let dateValue = new Date(dateObj.value);
                    if (dateValue >= FIRST_DAY_OF_YEAR && dateValue <= LAST_DAY_OF_YEAR) {
                        parsedDate.isDateComponentValid = true;
                        parsedDate.startDate = dateValue;
                        parsedDate.endDate = dateValue;
                    }
                }
            });
        }
        return parsedDate;
    }
}

module.exports.MyBot = LuisBot;
