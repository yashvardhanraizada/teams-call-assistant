// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using CallingBotSample.AdaptiveCards;
using CallingBotSample.Cache;
using CallingBotSample.Helpers;
using CallingBotSample.Models;
using CallingBotSample.Options;
using CallingBotSample.Services.MicrosoftGraph;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Text.Json;
using Newtonsoft.Json.Linq;
using static CallingBotSample.Bots.LLMClient;
using MeetingInfo = Microsoft.Graph.MeetingInfo;
using JsonSerializer = System.Text.Json.JsonSerializer;

namespace CallingBotSample.Bots
{
    public class MessageBot : TeamsActivityHandler
    {
        private readonly IAdaptiveCardFactory adaptiveCardFactory;
        private readonly AudioRecordingConstants audioRecordingConstants;

        private readonly ICallService callService;
        private readonly IChatService chatService;
        private readonly IOnlineMeetingService onlineMeetingService;
        private readonly IIncidentCache incidentCache;

        private readonly AzureAdOptions azureAdOptions;
        private readonly BotOptions botOptions;
        private readonly ILogger<MessageBot> logger;

        private readonly string _llmToken = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiI2OGRmNjZhNC1jYWQ5LTRiZmQtODcyYi1jNmRkZGUwMGQ2YjIiLCJpc3MiOiJodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGluZS5jb20vNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3L3YyLjAiLCJpYXQiOjE2OTQ2ODU3ODgsIm5iZiI6MTY5NDY4NTc4OCwiZXhwIjoxNjk0NjkwODI5LCJhaW8iOiJBYlFBUy84VUFBQUFrU05MRG5mbVkwaXl5Y2JaakJZOTJCaFJBeXZrWjFLZkcwdlRYZVJHNXVnVTl5NWg2ek15WjlPdms2UHZBVGpaTFRwWmdOTnc5QUhDb1F5RFFsVGdZekc2MzVBNWNrdXcwcmJHTUl2V1FzWFpocVVlREdwYkFacmpnQVllaVVEcTRJSW5QZkdQL21BNVhuVjV0MGNnR2pXdmJwK2twMEN6d0tBVlV4cmdZLzdxSnZka3g2WWhFUDNWR1RWQ25MZ2Ixb1NwT3lnYUdwdDZGTHhLQUYvNmw2b3ZnYXhTYmIxZ2I0c1lENEIyZXVVPSIsImF6cCI6IjY4ZGY2NmE0LWNhZDktNGJmZC04NzJiLWM2ZGRkZTAwZDZiMiIsImF6cGFjciI6IjAiLCJlbWFpbCI6ImtoeWF0aW9zd2FsQG1pY3Jvc29mdC5jb20iLCJuYW1lIjoiS2h5YXRpIE9zd2FsIiwib2lkIjoiYWViYjcxNDMtYWQ2YS00N2Y2LWEyMTMtMmEwNDZhOTY2MmZhIiwicHJlZmVycmVkX3VzZXJuYW1lIjoia2h5YXRpb3N3YWxAbWljcm9zb2Z0LmNvbSIsInJoIjoiMC5BUUVBdjRqNWN2R0dyMEdScXkxODBCSGJSNlJtMzJqWnl2MUxoeXZHM2Q0QTFySWFBQncuIiwic2NwIjoiYWNjZXNzIiwic3ViIjoiRUxCQ3RWMVFSd2o0cmwycGlhTTRiWEg1M3JIMWZoaTFTTURIRm8takRfbyIsInRpZCI6IjcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0NyIsInV0aSI6IkM3ODYyVHdRWWt1YlVmaUZ5SndIQUEiLCJ2ZXIiOiIyLjAiLCJ2ZXJpZmllZF9wcmltYXJ5X2VtYWlsIjpbImtoeWF0aW9zd2FsQG1pY3Jvc29mdC5jb20iXX0.L10N9-Kkv60SQELCjKqRJgvE-r3IydeDhTLzYoxnE_kO7LFQqlN3UCthZXbSorBS59AcfLyEpP5_7i5oMoGl88thVzAASvOVhWnMeB3ok-PeYxcDBGbKbmp3mpIhR1O6gYFTsbRhyPIRcMdknAbyweCWXY9gdnsR1uFdZgAIA-QnfHbhZFj7MjfsL0MtM7muf3yPKZFEUF949X8EjKXM_iLeVaMSQOjb7nlSzKr8D08YLIg8MyUkRqI14s8tyuPcOrCZivGhXUQXtuN6RUBERqrUrlYVEOzHlKRVeGXHew_iOEL28qlB66XmJU_w--V61Es8aP4ZfDRG1sWZ2Cbf0Q";
        public MessageBot(
            IAdaptiveCardFactory adaptiveCardFactory,
            AudioRecordingConstants audioRecordingConstants,
            ICallService callService,
            IChatService chatService,
            IOnlineMeetingService onlineMeetingService,
            IIncidentCache incidentCache,
            IOptions<AzureAdOptions> azureAdOptions,
            IOptions<BotOptions> botOptions,
            ILogger<MessageBot> logger)
        {
            this.adaptiveCardFactory = adaptiveCardFactory;
            this.audioRecordingConstants = audioRecordingConstants;

            this.callService = callService;
            this.chatService = chatService;
            this.onlineMeetingService = onlineMeetingService;
            this.incidentCache = incidentCache;

            this.azureAdOptions = azureAdOptions.Value;
            this.botOptions = botOptions.Value;
            this.logger = logger;
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(turnContext.Activity.Text))
            {
                dynamic value = turnContext.Activity.Value;
                if (value != null)
                {
                    string type = value["type"];
                    type = string.IsNullOrEmpty(type) ? "." : type.ToLower();
                    string? callId = value["callId"] ?? null;
                    await SendResponse(turnContext, type, callId, cancellationToken);
                }
            }
            else
            {
                turnContext.Activity.RemoveRecipientMention();
                await SendResponse(turnContext, turnContext.Activity.Text.Trim().ToLower(), null, cancellationToken);
            }
        }

        private async Task SendReplyAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken, string inputText)
        {
            //string summaryType = "Stacey's Response";
            string summarizedText = inputText;

            var card = new HeroCard { };
            //card.Title = summaryType;
            card.Text = summarizedText;

            var activity = MessageFactory.Attachment(card.ToAttachment());
            await turnContext.SendActivityAsync(activity, cancellationToken);
        }

        protected override Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var asJobject = JObject.FromObject(taskModuleRequest.Data);
            var fetchData = asJobject.ToObject<TaskModuleFetchData>();

            var taskInfo = new TaskModuleTaskInfo();

            switch (fetchData?.Action)
            {
                // Opens a module with a people picker where users can be selected. Later those user will be used to create a call
                case "createcall":
                    taskInfo.Card = adaptiveCardFactory.CreatePeoplePickerCard("Choose who to create a call with:", "Create", callId: null, isMultiSelect: true);
                    taskInfo.Title = "Create call";
                    break;
                // Opens a module with a people picker where a user can be selected to transfer the current call to
                case "transfercall":
                    taskInfo.Card = adaptiveCardFactory.CreatePeoplePickerCard("Choose who to transfer the call to:", "Transfer", fetchData?.CallId);
                    taskInfo.Title = "Transfer call";
                    break;
                // Opens a module with a people picker where a user can be selected to invite a participant to the current call
                case "inviteparticipant":
                    taskInfo.Card = adaptiveCardFactory.CreatePeoplePickerCard("Choose who to invite to the call:", "Invite", fetchData?.CallId);
                    taskInfo.Title = "Select the user to invite";
                    break;
                // Opens a modules with a form to create an incident. This includes a incident title, and those who should be on the call.
                case "openincidenttask":
                    taskInfo.Card = adaptiveCardFactory.CreateIncidentCard();
                    taskInfo.Title = "Create incident";
                    break;
                default:
                    break;
            }

            return Task.FromResult(new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse()
                {
                    Value = taskInfo,
                },
            });
        }

        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var asJobject = JObject.FromObject(taskModuleRequest.Data);
            var moduleSubmitData = asJobject.ToObject<TaskModuleSubmitData>();
            var peoplePicker = moduleSubmitData?.PeoplePicker;

            if (peoplePicker != null)
            {
                // Adaptive Card people picker returns a comma separated list of aad IDs
                var peoplePickerAadIds = peoplePicker.Split(',');
                var action = moduleSubmitData?.Action?.ToLowerInvariant();
                var callId = moduleSubmitData?.CallId;

                try
                {
                    switch (action)
                    {
                        case "create":
                            var call = await callService.Create(users: peoplePickerAadIds.Select(p => new Identity { Id = p }));

                            if (call != null)
                            {
                                await turnContext.SendActivityAsync(MessageFactory.Attachment(adaptiveCardFactory.CreateMeetingActionsCard(call.Id)));

                                return await CreateTaskModuleMessageResponse("Working on that, you can close this dialog now.");
                            }
                            break;
                        case "transfer":
                            return await CallService.HandleTeamsCallNotBeingFound(
                                callId,
                                (nonNullCallId) => callService.Transfer(
                                        nonNullCallId,
                                        new Identity { Id = peoplePicker }),
                                CreateTaskModuleMessageResponse);
                        case "invite":
                            return await CallService.HandleTeamsCallNotBeingFound(
                                        callId,
                                (nonNullCallId) => callService.InviteParticipant(
                                        nonNullCallId,
                                        new[] { new IdentitySet { User = new Identity { Id = peoplePicker } } }),
                                CreateTaskModuleMessageResponse);
                        case "createincident":
                            if (moduleSubmitData?.IncidentName != null)
                            {
                                return await CreateIncidentCall(
                                    turnContext,
                                    moduleSubmitData.IncidentName,
                                    peoplePickerAadIds,
                                    cancellationToken);
                            }
                            break;
                        default:
                            break;
                    }
                }
                catch (ServiceException ex)
                {
                    logger.LogError(ex, "Failure while making Graph Call");
                    return await CreateTaskModuleMessageResponse($"Something went wrong 😖. {ex.Message}");
                }
            }

            return await CreateTaskModuleMessageResponse("Something went wrong 😖");
        }

        private string ParseLLMResponseToString(string response)
        {
            LLMResponse responseObject = JsonConvert.DeserializeObject<LLMResponse>(response);
            return responseObject.Choices[0].Text;
        }

        private async Task<string> GetSummarizedText(string text)
        {
            string requestData = JsonSerializer.Serialize(new ModelPrompt
            {
                Prompt = text,
                MaxTokens = 500,
                Temperature = 0.6,
                TopP = 1,
                N = 1,
                Stream = false,
                LogProbs = null,
                Stop = ""
            });

            LLMClient llmClient = new LLMClient();
            var response = await llmClient.SendRequest("text-davinci-003", requestData, _llmToken);

            //Parse the responseText well and present it neatly.

            var summaryText = ParseLLMResponseToString(response);

            return summaryText;
        }

        private async Task SendResponse(ITurnContext<IMessageActivity> turnContext, string input, string? callId, CancellationToken cancellationToken)
        {
            if (input.Contains("settextcontext"))
            {
                MessageConstants.plainTextContext = input;
                await SendReplyAsync(turnContext, cancellationToken, "Text context is set. Proceed with your query either by text or call.");
                /*await turnContext.SendActivityAsync(
                        MessageFactory.Attachment(
                            adaptiveCardFactory.CreateInitiateCallCard()), cancellationToken);*/

                return;
            }
            else if (input.Contains("deletetextcontext"))
            {
                MessageConstants.plainTextContext = "";
                await SendReplyAsync(turnContext, cancellationToken, "Text context is deleted");
                return;
            }
            else if (input.Contains("setdocumentcontext"))
            {
                MessageConstants.documentContext = "";
                await SendReplyAsync(turnContext, cancellationToken, "Document context is set. Proceed with your query either by text or call.");
                return;
            }
            else if (input.Contains("setmeetingcontext"))
            {
                MessageConstants.meetingContext = "0:0:0.0 --> 0:0:0.130 Ketaki Ghotikar Yeah. 0:0:0.260 --> 0:0:1.670 Ketaki Ghotikar Can you go see my screen? 0:0:6.150 --> 0:0:6.330 Deepak R Yep. 0:0:7.760 --> 0:0:7.920 Ketaki Ghotikar Yeah. 0:1:24.710 --> 0:1:37.540 Ketaki Ghotikar So basically for this RT real-time monitoring alerting we have created three out-of-the-box rules right now for customers for each of the modalities. 0:1:37.550 --> 0:1:45.200 Ketaki Ghotikar So we want to monitor based on audio parameters, video parameters, and app sharing, which is nothing but screen sharing parameters. 0:1:45.750 --> 0:1:52.820 Ketaki Ghotikar So the requirement was that these went as straight different notifications for a particular user. 0:1:53.630 --> 0:1:59.540 Ketaki Ghotikar Ohh so the question was why not have a single rule for all three modalities? 0:2:0.430 --> 0:2:19.40 Ketaki Ghotikar So actually we had a very long discussion when this development or this feature was about to start where whether we should have three different rules or whether we should have a single rule adding all these monitoring settings and scope and everything per modality in the single rule. 0:2:19.980 --> 0:2:41.190 Ketaki Ghotikar Ohh so that time it was mainly decided to have three different rules for each of the modalities because of the reason that ohh the requirement was that for each modality we should send a different notification and then it should have different monitoring settings and parameters are anyways different. 0:2:42.80 --> 0:2:52.30 Ketaki Ghotikar So it would have been a really long rules page and maybe difficult also for an admin to understand all the parameters at a single go. 0:2:52.960 --> 0:2:59.840 Ketaki Ghotikar So that's why it was decided that we will keep them as separate three separate rules. 0:3:0.940 --> 0:3:2.910 Ketaki Ghotikar Ohm I you know. 0:3:2.950 --> 0:3:5.330 Ketaki Ghotikar So that was the decision at that time. 0:3:8.40 --> 0:3:17.50 Ketaki Ghotikar If we get the feedback from customers, otherwise we will have to think about it going forward if we can. 0:3:18.820 --> 0:3:28.270 Ketaki Ghotikar Basically, merge it into a single rule or still should have a single multiple rules, but probably combine the notification. 0:3:28.280 --> 0:3:33.310 Ketaki Ghotikar Maybe so we'll have to take it up based on customer feedback going forward. 0:3:38.570 --> 0:3:39.0 Ketaki Ghotikar Yeah. 0:7:51.330 --> 0:8:0.390 Ketaki Ghotikar Ohh yeah, So what happens is the question is with this we are binding ourselves to service fabric. 0:8:0.480 --> 0:8:3.150 Ketaki Ghotikar Is that supported as part of costing migration? 0:8:4.120 --> 0:8:7.170 Ketaki Ghotikar Is this making our service uh as stateful? 0:8:7.220 --> 0:8:9.370 Ketaki Ghotikar So the answer is yes. 0:8:10.60 --> 0:8:31.390 Ketaki Ghotikar So what happens is whenever for a user there is a meeting start event from Israel team, we create a service fabric so all our MTA applications are running in service fabric cluster and we are using some of the service fabric provider functionalities like actor reminder. 0:8:31.760 --> 0:8:40.160 Ketaki Ghotikar So we create an actor reminder ohm for this that particular user because it's real-time monitoring. 0:8:40.170 --> 0:8:41.700 Ketaki Ghotikar So we have to monitor. 0:8:41.810 --> 0:8:45.860 Ketaki Ghotikar So basically we need to check every X minutes for that user. 0:8:45.870 --> 0:8:54.900 Ketaki Ghotikar If the call quality in the last five minutes were good or not and based on that we have to send the notification if it is bad. 0:8:55.290 --> 0:8:58.960 Ketaki Ghotikar So that's why this is this has to be like recurring reminder. 0:8:59.710 --> 0:9:3.390 Ketaki Ghotikar So we are using service fabric actor reminder for that purpose. 0:9:4.190 --> 0:9:11.660 Ketaki Ghotikar Ohh so yes it is making us dependent on on the service fabric for this. 0:9:12.90 --> 0:9:18.600 Ketaki Ghotikar And yes, of course, if whenever cosmic migration has to be done, we have to get rid of this dependency. 0:9:18.810 --> 0:9:19.960 Ketaki Ghotikar And that was the major. 0:9:20.620 --> 0:9:26.310 Ketaki Ghotikar No Italia of prerequisites that we had for our cosmic migration. 0:9:27.110 --> 0:9:33.790 Ketaki Ghotikar Ohh which which had made this entire cosmic migration effort quite a lot. 0:9:34.530 --> 0:9:49.160 Ketaki Ghotikar So yeah, I mean we started using the same because in previous alerts also like device management we are using the same service fabric actor reminders features. 0:9:49.220 --> 0:9:53.790 Ketaki Ghotikar So we just have to have the same consistency. 0:9:53.800 --> 0:9:55.680 Ketaki Ghotikar We moved along the same lines. 0:9:56.770 --> 0:10:0.980 Ketaki Ghotikar Umm yeah, but for cosmic migration, we'll have to get rid of this. 0:10:0.990 --> 0:10:3.620 Ketaki Ghotikar And it does make the service as stateful. 0:16:16.450 --> 0:16:21.420 Ketaki Ghotikar So we have also have our RT dashboard on UI. 0:16:21.430 --> 0:16:28.190 Ketaki Ghotikar So basically what it shows is whenever you go to a user page and you see all the meetings. 0:16:28.640 --> 0:16:47.180 Ketaki Ghotikar So for that user, so some past or history meetings and some like recent or like ongoing meetings, so you can click on that ongoing meeting go to that page which is called RT dashboard and there you can see all the real-time telemetry for this user which includes all the. 0:16:47.240 --> 0:16:54.420 Ketaki Ghotikar So it's basically the same telemetry that we receive on the MTA event Hub, which is showcased on the RT dashboard. 0:16:54.760 --> 0:17:11.600 Ketaki Ghotikar But RT Dashboard is something that admin has to do manually, like subscribe to that meeting manually when the meeting is going on, and then only you'll be able to see the telemetry if he doesn't subscribe or whenever the meeting is going on, the telemetry will not be stored. 0:17:11.970 --> 0:17:18.940 Ketaki Ghotikar That is something manual and also admin has to keep monitoring that dashboard in order to know any about any issues. 0:17:19.190 --> 0:17:24.940 Ketaki Ghotikar So that's why the RT alerting, which is more proactive way of doing it, came in place. 0:17:25.790 --> 0:17:30.180 Ketaki Ghotikar\r\n";
                await SendReplyAsync(turnContext, cancellationToken, "Meeting context is set. Proceed with your query either by text or call.");
                return;
            }
            else if (input.Contains("answerme"))
            {
                MessageConstants.queryStatement = input;
                string completePrompt = MessageConstants.basePrompt + "\n\n" + MessageConstants.plainTextContext + "\n\n" + MessageConstants.documentContext + "\n\n" + MessageConstants.meetingContext + "\n\n" + MessageConstants.queryStatement;
                string response = await GetSummarizedText(completePrompt);
                await SendReplyAsync(turnContext, cancellationToken, response);
                MessageConstants.queryStatement = "";
                return;
            }

            switch (input)
            {
                case "playrecordprompt":
                    await CallService.HandleTeamsCallNotBeingFound(
                        callId,
                        (nonNullCallId) => callService.Record(nonNullCallId, audioRecordingConstants.PleaseRecordYourMessage),
                        (message) => UpdateActivityAsync(message, turnContext, cancellationToken));
                    break;
                case "hangup":
                    await CallService.HandleTeamsCallNotBeingFound(
                        callId,
                        (nonNullCallId) => callService.HangUp(nonNullCallId),
                        (message) => UpdateActivityAsync(message, turnContext, cancellationToken));
                    break;
                case "joinscheduledmeeting":
                    if (turnContext.Activity.ChannelData["meeting"] != null)
                    {
                        var call = await JoinScheduledMeeting(turnContext, cancellationToken);

                        if (call != null)
                        {
                            await turnContext.SendActivityAsync(MessageFactory.Attachment(adaptiveCardFactory.CreateMeetingActionsCard(call.Id)));
                        }
                    }
                    else
                    {
                        await turnContext.SendActivityAsync("Meeting not found. Are you calling this from a meeting chat?", cancellationToken: cancellationToken);
                    }
                    break;
                case "hi":
                    await turnContext.SendActivityAsync(
                        MessageFactory.Attachment(
                            adaptiveCardFactory.CreateWelcomeCard(turnContext.Activity.ChannelData["meeting"] != null)), cancellationToken);
                    break;
                default:
                    await SendReplyAsync(turnContext, cancellationToken, input);
                    break;
            }
        }

        private async Task<Call> JoinScheduledMeeting(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var organiser = await GetMeetingOrganiser(turnContext, cancellationToken);

            var channelDataTenant = JObject.Parse(JsonConvert.SerializeObject(turnContext.Activity.ChannelData)).SelectToken("tenant");
            organiser.SetTenantId(channelDataTenant["id"].ToString());

            return await callService.Create(
                new ChatInfo
                {
                    ThreadId = turnContext.Activity.Conversation.Id,
                    // NOTE: If you don't provide a Message Id, users will not be able to join the call the bot creates.
                    MessageId = "0"
                },
                new OrganizerMeetingInfo
                {
                    Organizer = new IdentitySet
                    {
                        User = organiser
                    },
                });
        }

        private async Task<Identity?> GetMeetingOrganiser(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var users = await TeamsInfo.GetPagedMembersAsync(turnContext, cancellationToken: cancellationToken);

            foreach (TeamsChannelAccount user in users.Members)
            {
                TeamsMeetingParticipant participant = await TeamsInfo.GetMeetingParticipantAsync(turnContext, participantId: user.AadObjectId).ConfigureAwait(false);

                if (participant.Meeting.Role == "Organizer")
                {
                    return new Identity
                    {
                        // This needs to be the organiser of the meeting, so you can't use the activity invoker
                        Id = user.AadObjectId,
                    };
                }
            }

            return null;
        }

        private async Task<TaskModuleResponse> CreateIncidentCall(ITurnContext turnContext, string incidentSubject, string[] peoplePickerAadIds, CancellationToken cancellationToken)
        {
            var onlineMeeting = await onlineMeetingService.Create(incidentSubject, peoplePickerAadIds);

            if (onlineMeeting != null)
            {
                MeetingInfo meetingInfo = JoinInfo.ParseMeetingInfo(onlineMeeting.JoinWebUrl);

                var meetingCall = await callService.Create(onlineMeeting.ChatInfo, meetingInfo);

                if (meetingCall != null)
                {
                    await chatService.InstallApp(meetingCall.ChatInfo.ThreadId, botOptions.CatalogAppId);

                    var incidentDetails = new IncidentDetails
                    {
                        CallId = meetingCall.Id,
                        IncidentSubject = incidentSubject,
                        MeetingInfo = meetingInfo,
                        ChatInfo = onlineMeeting.ChatInfo,
                        StartTime = DateTime.Now,
                        Participants = peoplePickerAadIds.Select(p => new Identity
                        {
                            Id = p,
                        })
                    };
                    incidentCache.Set(meetingCall.Id, incidentDetails);

                    await SendActivityToConversation(
                        turnContext,
                        onlineMeeting.ChatInfo.ThreadId,
                        MessageFactory.Attachment(adaptiveCardFactory.CreateIncidentMeetingCard(
                            incidentDetails.IncidentSubject,
                            incidentDetails.CallId,
                            incidentDetails.StartTime,
                            null
                        )),
                        cancellationToken);

                    await turnContext.SendActivityAsync("Created incident call successfully.", cancellationToken: cancellationToken);
                }

                return await CreateTaskModuleMessageResponse("Working on that, you can close this dialog now.");
            }

            return await CreateTaskModuleMessageResponse("Something went wrong 😖");
        }

        private async Task SendActivityToConversation(ITurnContext turnContext, string conversationId, IActivity activity, CancellationToken cancellationToken)
        {
            var newReference = new ConversationReference
            {
                Conversation = new ConversationAccount
                {
                    Id = conversationId
                },
                ServiceUrl = turnContext.Activity.ServiceUrl
            };

            await (turnContext.Adapter).ContinueConversationAsync(
                azureAdOptions.ClientId,
                newReference,
                async (ITurnContext turnContext, CancellationToken cancellationToken) =>
                {
                    await turnContext.SendActivityAsync(activity, cancellationToken);
                },
                cancellationToken);
        }

        private async Task<TaskModuleResponse> CreateTaskModuleMessageResponse(string value)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleMessageResponse()
                {
                    Value = value
                },
            };
        }

        private async Task UpdateActivityAsync(string responseText, ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var updatedActivity = MessageFactory.Text(responseText);
            updatedActivity.Id = turnContext.Activity.ReplyToId;
            await turnContext.UpdateActivityAsync(updatedActivity, cancellationToken);
        }
    }
}
