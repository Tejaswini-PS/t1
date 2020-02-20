// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// Generated with Bot Builder V4 SDK Template for Visual Studio EchoBot v4.6.2

using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using Demo12_DevBotAuth4EchoBot.LUIS;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Newtonsoft.Json;
using Microsoft.SharePoint.Client;//download 3 dll req and add as referance to sort pkg issue (look at assembly)
using AdaptiveCards;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Bot.Builder.Teams;
using System.Linq;
using System.Xml;

namespace Demo12_DevBotAuth4EchoBot.Bots
{
    public class EchoBottej1 : ActivityHandler
    {
        // Messages sent to the user.
        private const string WelcomeMessage = "Welcome to Bot demo which is combination of Auth+LUIS+SharePoint ";
        static string intentEntity = "";
        static string reasonintent = "";
        string greet = "";
        //Welcome card for users
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    string time = string.Format("{0:hh:mm:ss tt}", DateTime.Now);

                    if (time.Contains("PM"))
                        greet = "Hello!! How can I help You";
                    else
                        greet = "Hello!! How can I help You";
                    //await turnContext.SendActivityAsync(CreateActivityWithTextAndSpeak($" {greet} "), cancellationToken);
                }
            }
        }

        //Recive msg from users

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0));
            var items = new System.Collections.Generic.List<AdaptiveElement>();
            var items1 = new System.Collections.Generic.List<AdaptiveElement>();
            var items2 = new System.Collections.Generic.List<AdaptiveElement>();
            var items3 = new System.Collections.Generic.List<AdaptiveElement>();
            var items4 = new System.Collections.Generic.List<AdaptiveElement>();
            var items5 = new System.Collections.Generic.List<AdaptiveElement>();




            var text = turnContext.Activity.Text.ToLowerInvariant();
            var ans1 = text;
            //if (text.Contains("great") || text.Contains("fine") || text.Contains("awesome") || text.Contains("good") || text.Contains("ok"))
            //ans1 = "good";
            if (text.Equals("hello") || text.Equals("hi") || text.Equals("what's up"))
                ans1 = "hi";
            else if (text.Contains("how are you") || text.Contains("how do you do") || text.Contains("what about you") || text.Contains("what about you") || text.Contains("and you") || text.Contains("how you going"))
                ans1 = "how are you";
            // else if (text.Contains("great") || text.Contains("fine") || text.Contains("awesome") || text.Contains("good") || text.Contains("ok"))
            //   ans1 = "good";
            switch (ans1)
            {
                case "hi":
                    string time = string.Format("{0:hh:mm:ss tt}", DateTime.Now);

                    if (time.Equals("PM"))
                        await turnContext.SendActivityAsync($"Hello!!, How are you doing?", cancellationToken: cancellationToken);
                    else
                        await turnContext.SendActivityAsync($"Hello!!, How are you doing?", cancellationToken: cancellationToken);

                    break;
                case "how are you":
                    await turnContext.SendActivityAsync($"I am fine thank you, How can I help you?", cancellationToken: cancellationToken);
                    break;
                case "good":
                    await turnContext.SendActivityAsync($"How can I help you?", cancellationToken: cancellationToken);
                    break;
                case "intro":
                    await turnContext.SendActivityAsync($"Here is R2D2, I am 24X7 available for you, I am helping to solve all UAT application queries,what can i help you?", cancellationToken: cancellationToken);

                    break;
                //case "help":
                //    await SendIntroCardAsync(turnContext, cancellationToken);
                //    break;
                default:
                    List<Leave> list1 = new List<Leave>();

                    list1 = await LuisDatabaSourceCalls(turnContext.Activity.Text);
                    if (reasonintent.Contains("leave") || reasonintent.Equals("leave") || reasonintent == "leave")
                    {


                        int a = list1.Count;


                        for (int i = 0; i < list1.Count; i++)
                        {

                            items.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].Team,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent
                            });
                        }

                        for (int i = 0; i < list1.Count; i++)
                        {

                            items1.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].StartDate,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent
                            });
                        }
                        for (int i = 0; i < list1.Count; i++)
                        {

                            items2.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].Author,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent

                            });
                        }
                        for (int i = 0; i < list1.Count; i++)
                        {

                            items3.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].Reason,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent
                            });
                        }


                        //**** drop down in menu 

                        System.Uri uri = new System.Uri("C:/Users/tejas/OneDrive/Desktop/DevBotAuth4EchoBot_24_1_2020/Demo12_DevBotAuth4EchoBot/Demo12_DevBotAuth4EchoBot/Images/Testimage.png");
                        // string strtej = "C:\Users\tejas\OneDrive\Desktop\DevBotAuth4EchoBot_24_1_2020\Demo12_DevBotAuth4EchoBot\Demo12_DevBotAuth4EchoBot\Images\Testimage.png";"
                        card.Body.Add(new AdaptiveContainer()
                        {


                            Items = new List<AdaptiveElement>()
                    {


                            new AdaptiveTextBlock()
                            {

                                Type="TextBlock",
                               Text="Masycoda",
                               Height=AdaptiveHeight.Stretch,
                               Color=AdaptiveTextColor.Good,
                               Size=AdaptiveTextSize.Large,
                               Weight=AdaptiveTextWeight.Bolder,
                               HorizontalAlignment=AdaptiveHorizontalAlignment.Right,


                            },


                             new AdaptiveTextBlock()
                            {



                                 Type="TextBlock",
                               Text="Pvt Ltd",

                               HorizontalAlignment=AdaptiveHorizontalAlignment.Right,


                            },

                          new AdaptiveColumnSet()
                            {
                                Type = "ColumnSet",
                                Height = AdaptiveHeight.Auto,
                                Columns=new List<AdaptiveColumn> ()
                                {
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="100px",

                                      Items=new List<AdaptiveElement>()
                                      {

                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Name",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color=AdaptiveTextColor.Warning
                                          }
                                      }
                                  },
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="70px",
                                      Items=new List<AdaptiveElement>()
                                      {
                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Team",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color = AdaptiveTextColor.Warning

                                          }
                                      }
                                  },
                                    new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="100px",
                                      Items=new List<AdaptiveElement>()
                                      {
                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Reason",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color =AdaptiveTextColor.Warning
                                          }
                                      }
                                  },
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="125px",
                                      Items=new List<AdaptiveElement>()
                                      {
                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Date",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color =AdaptiveTextColor.Warning
                                               //Width= "123px";
                                          }
                                      }
                                  }






                                }
                          },


                        new AdaptiveColumnSet()
                        {
                            Type = "ColumnSet",
                            Height = AdaptiveHeight.Auto,

                            Columns=new List<AdaptiveColumn> ()
                            {

                                  new AdaptiveColumn()
                              {
                                  Type="Column",
                                  Width="100px",
                                 Items= items2,
                              }
                               ,
                              new AdaptiveColumn()
                              {
                                  Type="Column",
                                  Width="70px",
                                  Items=items,


                              },

                               new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="110px",
                                     Items= items3,
                                  },
                              new AdaptiveColumn()
                              {
                                  Type="Column",
                                  Width="125px",
                                 Items= items1,
                              }


                            }


                        },
                    }
                        });


                    }
                    if (reasonintent.Contains("reason") || reasonintent.Equals("reason") || reasonintent == "reason")
                    {

                        int a = list1.Count;



                        for (int i = 0; i < list1.Count; i++)
                        {

                            items2.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].Author,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent

                            });
                        }
                        for (int i = 0; i < list1.Count; i++)
                        {
                            items3.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].Reason,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent

                            });
                        }



                        System.Uri uri = new System.Uri("C:/Users/tejas/OneDrive/Desktop/DevBotAuth4EchoBot_24_1_2020/Demo12_DevBotAuth4EchoBot/Demo12_DevBotAuth4EchoBot/Images/Testimage.png");
                        card.Body.Add(new AdaptiveContainer()
                        {

                            Items = new List<AdaptiveElement>()
                    {


                            new AdaptiveTextBlock()
                            {

                                Type="TextBlock",
                               Text="Masycoda",
                               Height=AdaptiveHeight.Stretch,
                               Color=AdaptiveTextColor.Good,
                               Size=AdaptiveTextSize.Large,
                               Weight=AdaptiveTextWeight.Bolder,
                               HorizontalAlignment=AdaptiveHorizontalAlignment.Right,


                            },


                             new AdaptiveTextBlock()
                            {



                                 Type="TextBlock",
                               Text="Pvt Ltd",

                               HorizontalAlignment=AdaptiveHorizontalAlignment.Right,


                            },

                          new AdaptiveColumnSet()
                            {
                                Type = "ColumnSet",
                                Height = AdaptiveHeight.Auto,
                                Columns=new List<AdaptiveColumn> ()
                                {
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="200px",

                                      Items=new List<AdaptiveElement>()
                                      {

                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Name",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color=AdaptiveTextColor.Warning

                                          }
                                      }
                                  },

                                    new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="200px",
                                      Items=new List<AdaptiveElement>()
                                      {
                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Reason",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color =AdaptiveTextColor.Warning
                                          }
                                      }
                                  }






                                }
                          },


                        new AdaptiveColumnSet()
                        {
                            Type = "ColumnSet",
                            Height = AdaptiveHeight.Auto,

                            Columns=new List<AdaptiveColumn> ()
                            {

                                  new AdaptiveColumn()
                              {
                                  Type="Column",
                                  Width="200px",

                                 Items= items2,
                              }
                               ,


                               new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="200px",


                                     Items= items3,
                                  }



                            }


                        },
                    }
                        });


                    }
                    // string userAns = await LuisDatabaSourceCalls(turnContext.Activity.Text);

                    await turnContext.SendActivityAsync(CreateActivityWithTextAndSpeak($"LUIS Key: {intentEntity}"), cancellationToken);
                    //  await turnContext.SendActivityAsync(CreateActivityWithTextAndSpeak($" {userAns}"), cancellationToken);
                    if (intentEntity.Contains("start") && intentEntity.Contains("uat"))
                    {
                        await turnContext.SendActivityAsync(CreateActivityWithTextAndSpeak($"Please let me know your role in your organization 1-Super Admin ,2-Lead,3-Stakeholder,4-Test Pass Manager,5-Tester"), cancellationToken);
                    }

                    var attachment = new Microsoft.Bot.Schema.Attachment
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = card,
                    };
                    var reply = MessageFactory.Attachment(attachment);
                    await turnContext.SendActivityAsync(reply, cancellationToken);

                    break;
            }


            // Save any state changes.
            // await _userState.SaveChangesAsync(turnContext);
        }

        //LUIS
        private static async Task<List<Leave>> LuisDatabaSourceCalls(string Query)
        {

            string UserAns = "";
            UAT_LUIS_Entity UAT_LUIS = await GetEntityFromLUIS(Query);
            List<Leave> list1 = new List<Leave>();


            //UserAns = await GetDynamicQueryForSP("leave","today");

            //UserAns = await GetAnsDetailsFromSharepoint("add", "project");
            if (!UAT_LUIS.intents.Equals(null))
            {
                if (UAT_LUIS.intents.Length > 0 && UAT_LUIS.entities.Length > 0)
                {
                    string priorEntity = "";
                    //for app123
                    //for (int i = 0; i < UAT_LUIS.entities.Length; i++)
                    //{
                    //    if (UAT_LUIS.entities[i].role != null)
                    //    {
                    //        if (UAT_LUIS.entities[i].role.ToString().ToLower() == "origin")
                    //        {
                    //            priorEntity = UAT_LUIS.entities[i].entity.ToString();
                    //            break;
                    //        }
                    //    }

                    //}
                    //if (priorEntity == "")
                    priorEntity = UAT_LUIS.entities[0].type.ToString();
                    intentEntity = "Intent: " + UAT_LUIS.intents[0].intent.ToString() + "  Entity: " + priorEntity;
                    reasonintent = UAT_LUIS.intents[0].intent.ToString();
                    // UserAns = await GetAnsDetailsFromSharepoint(UAT_LUIS.intents[0].intent.ToString().ToLower(), priorEntity.ToLower());
                    //   list1 = await GetDynamicQueryForSP(UAT_LUIS.intents[0].intent.ToString().ToLower(), priorEntity.ToLower());
                    list1 = await GetDynamicQueryForSP(UAT_LUIS);

                    //  UserAns = await GetDynamicQueryForSP(UAT_LUIS);//(UAT_LUIS.intents[0].intent.ToString().ToLower(), priorEntity.ToLower());

                }
                else
                {
                    intentEntity = "";
                    UserAns = "Sorry, I am not getting you...No Intent or Entity Found....Could you please rephrase your query";

                }
            }
            else
                UserAns = "Sorry, LUIS Entity and Intent are null";

            return list1;
        }
        private static async Task<UAT_LUIS_Entity> GetEntityFromLUIS(string Query)
        {
            Query = Uri.EscapeDataString(Query);
            UAT_LUIS_Entity Data = new UAT_LUIS_Entity();
            using (HttpClient client = new HttpClient())
            {

                string LuisAppID = "11838aa9-ca07-4c63-9206-7e422f3a50bd";//"3d87a0d1-a41c-4c2c-b48b-a9bd8c308e61";//"5e376704-da8b-406e-9c68-a7ec15ea29eb";//"6473c60c-64d4-4b4e-895c-946deb9f27bd";//"0e4f70af-c3af-414b-a28f-e9861b17d3ab";//"6473c60c-64d4-4b4e-895c-946deb9f27bd";//"6f963293-9c7a-4393-87eb-9d9a50459662";//ConfigurationManager.AppSettings["aad:LuisAppId"];//["aad:LuisAppId"];
                string LuisSecretKey = "f40116b88dc24df9812735f71ab570e3";//"05b8b88ba1b7480388969d921f5d908c";//"e7301392690a40e387f2a300899d9a4d";//"9f0672f0109247928a0144ca4a910219";//"619b66fbeed048e586e81793f5a69925";//ConfigurationManager.AppSettings["aad:LuisSecretKey"];
                string RequestURI = "https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/" + LuisAppID + "?verbose=true&timezoneOffset=0&subscription-key=" + LuisSecretKey + "&q=" + Query;

                HttpResponseMessage msg = await client.GetAsync(RequestURI);

                if (msg.IsSuccessStatusCode)
                {
                    var JsonDataResponse = await msg.Content.ReadAsStringAsync();
                    Data = JsonConvert.DeserializeObject<UAT_LUIS_Entity>(JsonDataResponse);
                }
            }
            return Data;
        }

        //END

        //DB
        private static async Task<List<Leave>> GetDynamicQueryForSP(UAT_LUIS_Entity UAT_LUIS)//(string QueryIntent, string QueryEntity)
        {

            List<Leave> list1 = new List<Leave>();
            string leaveStartDate = "", leaveEndDate = "";
            string manager = "", PName = "", query = "", labelToMakeQuery = "";
            string userAns = "";
            List<string> SpListName = new List<string>();
            SpListName.Add("LeaveManagement");
            SpListName.Add("Attendance");
            //if (UAT_LUIS.intents[0].intent.ToString().ToLower().Equals("leave"))
            //{
            List<string> entityList = new List<string>();
            string finalEntity = UAT_LUIS.entities[0].type.ToString().ToLower();
            int noOfDuration = 0;
            for (int i = 0; i < UAT_LUIS.entities.Length; i++)
            {
                entityList.Add(UAT_LUIS.entities[i].type.ToString().ToLower());
            }
            if (entityList.Contains("last") == true)
            {
                if (entityList.Contains("days") == true)
                    finalEntity = "last days";
                else if (entityList.Contains("week") == true)
                    finalEntity = "last week";
                else if (entityList.Contains("month") == true)
                    finalEntity = "last month";
                else if (entityList.Contains("year") == true)
                    finalEntity = "last year";

            }
            else if (entityList.Contains("upcoming") == true)
            {
                if (entityList.Contains("days") == true)
                    finalEntity = "next days";
                else if (entityList.Contains("week") == true)
                    finalEntity = "next week";
                else if (entityList.Contains("month") == true)
                    finalEntity = "next month";
                else if (entityList.Contains("year") == true)
                    finalEntity = "next year";
            }
            else if (entityList.Contains("current") == true)
            {
                if (entityList.Contains("days") == true)
                    finalEntity = "current days";
                else if (entityList.Contains("week") == true)
                    finalEntity = "current week";
                else if (entityList.Contains("month") == true)
                    finalEntity = "current month";
                else if (entityList.Contains("year") == true)
                    finalEntity = "current year";
            }
            else if (entityList.Contains("today") == true)
            {
                finalEntity = "today";
            }
            else if (entityList.Contains("tomorrow") == true)
            {
                finalEntity = "tomorrow";
            }
            else if (entityList.Contains("yesterday") == true)
            {
                finalEntity = "yesterday";
            }
            if (entityList.Contains("builtin.number") == true)
            {
                int index = entityList.FindIndex(a => a.Contains("builtin.number"));
                var firstItem = (UAT_LUIS.entities[index].entity.ToString());

                switch (firstItem.ToLower())
                {
                    case "one":
                        firstItem = (1).ToString();
                        break;
                    case "two":
                        firstItem = (2).ToString();
                        break;
                    case "three":
                        firstItem = (3).ToString();
                        break;
                    case "four":
                        firstItem = (4).ToString();

                        break;
                    case "five":
                        firstItem = (5).ToString();
                        break;
                    case "six":
                        firstItem = (6).ToString();
                        break;
                    case "seven":
                        firstItem = (7).ToString();
                        break;
                    case "eight":
                        firstItem = (8).ToString();
                        break;
                    case "nine":
                        firstItem = (9).ToString();
                        break;
                    case "ten":
                        firstItem = (10).ToString();
                        break;
                }

                //  String ast = entityList.FindIndex(a => a.Contains("builtin.number")); ;
                //mychanges  int.TryParse(UAT_LUIS.entities[index].entity.ToString(), out noOfDuration);
                int.TryParse(firstItem, out noOfDuration);
            }
            if (entityList.Contains("date") == true)
            {

                var indexVal = Enumerable.Range(0, entityList.Count)
                             .Where(i => entityList[i] == "date")
                             .ToList();
                for (int i = 0; i < indexVal.Count; i++)
                {
                    if (UAT_LUIS.entities[indexVal[i]].role.ToString().Equals("FromDate"))
                        leaveStartDate = Convert.ToDateTime(leaveStartDate).ToString("yyyy-MM-ddTHH:mm:ssZ");
                    else if (UAT_LUIS.entities[i].role.ToString().Equals("ToDate"))
                        leaveEndDate = Convert.ToDateTime(leaveEndDate).ToString("yyyy-MM-ddTHH:mm:ssZ");




                }
                // int.TryParse(UAT_LUIS.entities[index].entity.ToString(), out noOfDuration);
            }
            switch (finalEntity)
            {
                case "today":
                    {
                        leaveStartDate = leaveEndDate = DateTime.UtcNow.ToString("yyyy-MM-ddT18:30:00Z");
                        labelToMakeQuery = "dateRange";
                        break;
                    }
                case "tomorrow":
                    {
                        leaveStartDate = leaveEndDate = DateTime.UtcNow.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";
                        break;
                    }
                case "yesterday":
                    {
                        leaveStartDate = leaveEndDate = DateTime.UtcNow.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";
                        break;
                    }
                case "the day after tomorrow":
                    {
                        //case "the day after tomorrow":
                        leaveStartDate = DateTime.UtcNow.AddDays(2).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        leaveEndDate = DateTime.UtcNow.AddDays(2).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";
                        break;
                    }
                case "the day before yesterday":
                    {
                        // case "the day before yesterday":
                        leaveStartDate = DateTime.UtcNow.AddDays(-2).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        leaveEndDate = DateTime.UtcNow.AddDays(-2).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";
                        break;
                    }
                case "upcoming days":
                    {
                        //we consider here for 5 days
                        //  case "upcoming days":
                        leaveStartDate = DateTime.UtcNow.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        leaveEndDate = DateTime.UtcNow.AddDays(5).ToString("yyyy-MM-ddTHH:mm:ssZ");

                        labelToMakeQuery = "dateRange";
                        break;
                    }
                case "current week":
                    {

                        //  case "current week":
                        leaveStartDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        leaveEndDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddDays(7).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";
                        break;
                    }
                case "last week":
                    { //case "last week":
                        leaveStartDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddDays(-7 * noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        leaveEndDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";
                        break;
                    }
                case "next week":
                    {
                        //  case "next week":
                        if (noOfDuration != 0)
                        {
                            leaveStartDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddDays(7).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddDays(noOfDuration * 7 + 7).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        break;
                    }
                case "last days":
                    {

                        if (noOfDuration != 0)
                        {
                            //   case "last days":
                            leaveStartDate = DateTime.UtcNow.AddDays(-noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = DateTime.UtcNow.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        else
                            labelToMakeQuery = "dateNotFound";
                        break;
                    }
                case "next days":
                    {
                        // int no = 0;
                        //int.TryParse(UAT_LUIS.entities[0].entity.ToString(), out no);
                        if (noOfDuration != 0)
                        {
                            //case "next days":
                            leaveStartDate = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = DateTime.UtcNow.AddDays(noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");

                            labelToMakeQuery = "dateRange";
                        }
                        else
                            labelToMakeQuery = "dateNotFound";
                        break;
                    }
                case "last month":
                    {
                        //int no = 0;
                        //int.TryParse(UAT_LUIS.entities[0].entity.ToString(), out no);
                        if (noOfDuration != 0)
                        {
                            //  case "last month":
                            leaveStartDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(-noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        else
                            labelToMakeQuery = "dateNotFound";
                        break;

                    }
                case "next month":
                    {
                        ///chk
                        //int no = 0;
                        //int.TryParse(UAT_LUIS.entities[0].entity.ToString(), out no);
                        if (noOfDuration != 0)
                        {
                            // case "next month":
                            leaveStartDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        else
                            labelToMakeQuery = "dateNotFound";
                        break;
                    }
                case "current month":
                    {
                        //case "current month":
                        leaveStartDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        leaveEndDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(1).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";

                        break;
                    }
                case "last year":
                    {

                        if (noOfDuration != 0)
                        {
                            //  case "last year":
                            leaveStartDate = new DateTime(DateTime.UtcNow.AddYears(-noOfDuration).Year, 1, 1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddYears(-noOfDuration).Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        else
                            labelToMakeQuery = "dateNotFound";
                        break;
                    }
                case "next year":
                    {
                        if (noOfDuration != 0)
                        {
                            // case "next year":
                            leaveStartDate = new DateTime(DateTime.UtcNow.AddYears(noOfDuration).Year, 1, 1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddYears(noOfDuration).Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        else
                            labelToMakeQuery = "dateNotFound";
                        break;
                    }
                case "current year":
                    { //  case "current year":
                        leaveStartDate = new DateTime(DateTime.UtcNow.Year, 1, 1).ToString("yyyy-MM-ddTHH:mm:ssZ");//DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                        leaveEndDate = new DateTime(DateTime.UtcNow.Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";

                        break;
                    }

                case "query":
                    {
                        break;
                    }
            }
            if (entityList.Contains("manager") == true)
            {
                int index = entityList.FindIndex(a => a.Contains("manager"));
                manager = UAT_LUIS.entities[index].entity.ToString();
                labelToMakeQuery = "dateRangeWithManager";
            }
            if (entityList.Contains("pname") == true)
            {
                int index = entityList.FindIndex(a => a.Contains("pname"));
                PName = UAT_LUIS.entities[index].entity.ToString();
                labelToMakeQuery = "reasonForPName";
            }
            //}
            switch (UAT_LUIS.intents[0].intent.ToString())
            {
                case "leave":
                    switch (labelToMakeQuery)
                    {
                        case "dateRange":
                            {
                                //    query = string.Format("<View><Query><ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='Title' /></ViewFields><Where><And>" +
                                //"<Geq><FieldRef Name='StartDate' /><Value Type='DateTime'>2020-02-06T18:30:00Z</Value></Geq>" +
                                //"<Leq><FieldRef Name='EndDate' /><Value Type='DateTime'>{1}</Value></Leq>" +
                                //"</And></Where></Query></View>", leaveStartDate, leaveEndDate);
                                query = string.Format("<View><Query><ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='Title' /></ViewFields><Where><And>" +
                                "<Geq><FieldRef Name='StartDate' /><Value Type='DateTime'>{0}</Value></Geq>" +
                                "<Leq><FieldRef Name='EndDate' /><Value Type='DateTime'>{1}</Value></Leq>" +
                                "</And></Where></Query></View>", leaveStartDate, leaveEndDate);
                                //TM userAns = await GetAnsDetailsFromSharepoint(query,SpListName[0]);
                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[0]);
                                //if (userAns == "")
                                //    userAns = "No data found for specific date";
                                break;
                            }
                        case "dateRangeWithManager":
                            {
                                query = string.Format("<View><Query><Where><And>" +
                                    "<Geq><FieldRef Name='StartDate' /><Value  Type='DateTime'>{0}</Value></Geq>" +
                                    "<And><Leq><FieldRef Name='EndDate' /><Value  Type='DateTime'>{1}</Value></Leq>" +
                                    "<Eq><FieldRef Name='Manager' /><Value Type='UserMulti'>{2}</Value></Eq>" +
                                    "</And></And></Where></Query></View>", leaveStartDate, leaveEndDate, manager);
                                //TM  userAns = await GetAnsDetailsFromSharepoint(query, SpListName[0]);

                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[0]);
                                if (userAns == "")
                                    userAns = "No data found for specific date";
                                break;
                            }
                    }
                    break;

                case "reason":
                    //switch (labelToMakeQuery)
                    //{
                    //    case "reasonForPName":
                    //        {

                    //            //query = string.Format("<View><Query><Where><And>" +
                    //            //    "<Geq><FieldRef Name='StartDate' /><Value  Type='DateTime'>{0}</Value></Geq>" +
                    //            //    "<And><Leq><FieldRef Name='EndDate' /><Value  Type='DateTime'>{1}</Value></Leq>" +
                    //            //    "<Contains><FieldRef Name='Author' /><Value Type='User'>{2}</Value></Contains>" +
                    //            //    "</And></And></Where></Query></View>", leaveStartDate, leaveEndDate, PName);


                    //            string.Format("<View><Query></Query></View>");
                    //            list1 = await GetAnsDetailsFromSharepoint(query, SpListName[1]);


                    //            //    userAns = "No data found for specific date";
                    //            break;
                    //        }


                    //}
                    query = string.Format("<View><Query></Query></View>");
                    list1 = await GetAnsDetailsFromSharepoint(query, SpListName[0]);

                    break;
                case "WorkFromHome":

                    break;
            }
            //return userAns;

            return list1;
        }

        private static async Task<List<Leave>> GetAnsDetailsFromSharepoint(string Query, string listName)
        {
            //***********************
            string text = "";
            List<Leave> list1 = new List<Leave>();
            using (ClientContext clientContext = new ClientContext("https://masycodasolutions.sharepoint.com/sites/OfficeMgmt"))
            {

                SecureString password = new SecureString();
                string pass = "chhaya@123";
                foreach (char c in pass.ToCharArray()) password.AppendChar(c);
                clientContext.Credentials = new SharePointOnlineCredentials("krutika@masycoda.com", pass);

                // TimeZoneInfo.ConvertTimeFromUtc(dt, TimeZoneInfo);
                List UatList_ls = clientContext.Web.Lists.GetByTitle(listName);
                //camlquery query = camlquery.createallitemsquery(100);
                //List UatList_ls = clientContext.Web.Lists.GetByTitle("demoList");

                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = Query;
                //camlQuery.ViewXml = string.Format("<View><Query></Query></View>");
                ListItemCollection items = UatList_ls.GetItems(camlQuery);
                clientContext.Load(items);
                await clientContext.ExecuteQueryAsync();

                //** Person Column 
                //for (int i = 0; i < items.Count; i++)
                //{
                //    var childIdField = items[i]["Editor"] as FieldLookupValue;
                //    var childId_Value = childIdField.LookupValue;
                //    text += "\n" + items[i]["StartDate"].ToString();
                //    //var childIdField = items[i]["Manager"] as FieldLookupValue; 


                //}

                for (int i = 0; i < items.Count; i++)
                {
                    //***Added first time
                    // text += "\n" + items[i]["StartDate"].ToString();
                    //****

                    var childIdField = items[i]["Author"] as FieldLookupValue;
                    var childId_Value = childIdField.LookupValue;

                    var childIdField1 = items[i]["Manager"] as FieldLookupValue;
                    var childId_Value1 = childIdField.LookupValue;
                    String str1 = "", str2 = "";
                    if (items[i]["Team"].ToString() != null)
                    {
                        str1 = items[i]["Team"].ToString();
                    }

                    if (items[i]["LeaveType"].ToString() != null)
                    {
                        str2 = items[i]["LeaveType"].ToString();
                    }
                    //Changed String date1 = String.Concat(Convert.ToDateTime(items[i]["StartDate"]).AddHours(6).ToString("dd/MM/yyyy"), Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd/MM/yyyy"));
                    //  String date2= String.Concat(date1, Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd/MM/yyyy"));
                    String date1 = String.Concat(Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd/MM/yyyy"), " ");
                    String date2 = String.Concat(date1, Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd/MM/yyyy"));
                    // Reason = items[i]["LeaveType"].ToString(),

                    list1.Add(new Leave()
                    {
                        Author = childId_Value,
                        // StartDate = items[i]["StartDate"].ToString(),
                        //    StartDate = Convert.ToDateTime(items[i]["StartDate"]).AddHours(6).ToString("dd-MM-yyyy"),
                        //  EndDate = Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd-MM-yyyy"),
                        Reason = items[i]["LeaveType"].ToString(),
                        StartDate = date2,
                        Team = str1,
                        Manager = childId_Value1,
                    });


                }




                //*******************************


            }
            // return text;

            return list1;
        }

        //DB End



        //support speach also
        private IActivity CreateActivityWithTextAndSpeak(string message)
        {
            var activity = MessageFactory.Text(message);
            string speak = @"<speak version='1.0' xmlns='https://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
              <voice name='Microsoft Server Speech Text to Speech Voice (en-US, JessaRUS)'>" +
              $"{message}" + "</voice></speak>";
            activity.Speak = speak;
            return activity;
        }
    }
}
//______________________________________________________________________________________________________________________________________


