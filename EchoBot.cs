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
using System.Globalization;
using Microsoft.Extensions;
using GraphTutorial;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using System.Net;

namespace Demo12_DevBotAuth4EchoBot.Bots
{
    public class EchoBot : ActivityHandler
    {
        // Messages sent to the user.
        private const string WelcomeMessage = "Welcome to Bot demo which is combination of Auth+LUIS+SharePoint ";
        static string intentEntity = "";
        string greet = "";
        static string reasonintent = "";
        static List<string> SpListName = new List<string>();
        static string staticListName;
        static int countOfLeave = 0;
        static string starttimecalender, endtimecalender;
        string botanswer, dateluis, botanswer1;
        String new1, new2, new3, new4, newDateTime1;
        String roomluisname = "";
        DateTime dateresultoverlap, datevar;

        String ResultFromTime, ResultToTime, ResultFromtimeHour, ResultFromTimemMin, ResultToTimeHour, ResultToTimeMin;
        private IEnumerable<string> scopes = new string[] { "User.Read.All", "Calendars.ReadWrite" };

        // static string userInput = "";
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

                }
            }
        }






        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {





            ///***************************** changed  for to time and from time***********************
            ///


            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
            .Create("e34dfac1-8d12-45ce-99dc-ee97478abc48")
            .WithTenantId("fbcab41b-0c15-41f2-9858-b64924a83a6c").WithRedirectUri("http://localhost")
            .Build();


            //DateTime t = DateTime.Today;
            //DateTime t1 = t.AddHours(00);
            //DateTime t23 = t1.AddMinutes(00);
            //String s = Convert.ToString(t23);
            //string newDateTime = "";


            //DateTime dt;
            //if (DateTime.TryParse(t23.ToString(), out dt))
            //{
            //    newDateTime = dt.ToString("yyyy-MM-ddTHH:mm:ss.fffffffK");
            //}

            //String starttimeluis = newDateTime;




            var password = new SecureString();

            password.AppendChar('<');
            password.AppendChar('T');
            password.AppendChar('N');
            password.AppendChar('>');
            password.AppendChar('7');
            password.AppendChar('9');
            password.AppendChar('M');
            password.AppendChar('A');
            password.AppendChar('S');
            password.AppendChar('Y');
            password.AppendChar('C');
            password.AppendChar('O');
            password.AppendChar('D');
            password.AppendChar('A');



            UsernamePasswordProvider authProvider = new UsernamePasswordProvider(publicClientApplication, scopes);
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            Microsoft.Graph.User me = graphClient.Me.Request()
                .WithUsernamePassword("Tejaswini@indica.onmicrosoft.com", password)
                .GetAsync().Result;
            var eventss = (await graphClient.Me.Events
  .Request()
  .Header("Prefer", "outlook.timezone=\"UTC\"")
  .Select(e => new
  {
      e.Subject,
      e.Body,
      e.BodyPreview,
      e.Organizer,
      e.Attendees,
      e.Start,
      e.End,
      e.Location
  })
  .GetAsync()).ToList();











            //********************










            //***************************
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2));
            var items = new System.Collections.Generic.List<AdaptiveElement>();
            var items1 = new System.Collections.Generic.List<AdaptiveElement>();
            var items2 = new System.Collections.Generic.List<AdaptiveElement>();
            var items3 = new System.Collections.Generic.List<AdaptiveElement>();
            var items4 = new System.Collections.Generic.List<AdaptiveElement>();
            var items5 = new System.Collections.Generic.List<AdaptiveElement>();

            var dateAdaptive = new System.Collections.Generic.List<AdaptiveElement>();
            var WorkingLocationAdaptive = new System.Collections.Generic.List<AdaptiveElement>();
            var PresentAdaptive = new System.Collections.Generic.List<AdaptiveElement>();
            var AssignedToAdaptive = new System.Collections.Generic.List<AdaptiveElement>();


            //******how many event cards
            var FindRoom = new System.Collections.Generic.List<AdaptiveElement>();
            var Eventsname1 = new System.Collections.Generic.List<AdaptiveElement>();
            var EventDate = new System.Collections.Generic.List<AdaptiveElement>();
            var EventLocation = new System.Collections.Generic.List<AdaptiveElement>();

            ////***********


            var EventNameschedule = new System.Collections.Generic.List<AdaptiveElement>();
            var Eventschedulestart = new System.Collections.Generic.List<AdaptiveElement>();
            var Eventscheduleend = new System.Collections.Generic.List<AdaptiveElement>();
            var Eventschedulelocation = new System.Collections.Generic.List<AdaptiveElement>();
            var Eventschedulesubject = new System.Collections.Generic.List<AdaptiveElement>();


            ////***********

           





            //********


            var text = turnContext.Activity.Text.ToLowerInvariant();
            var ans1 = text;

            if (text.Equals("hello") || text.Equals("hi") || text.Equals("what's up"))
                ans1 = "hi";
            else if (text.Contains("how are you") || text.Contains("how do you do") || text.Contains("what about you") || text.Contains("what about you") || text.Contains("and you") || text.Contains("how you going"))
                ans1 = "how are you";

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
                    if (reasonintent.Equals("leave") || reasonintent == "leave")
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




                        System.Uri uri = new System.Uri("C:/Users/tejas/OneDrive/Desktop/DevBotAuth4EchoBot_24_1_2020/Demo12_DevBotAuth4EchoBot/Demo12_DevBotAuth4EchoBot/Images/Testimage.png");
                        // string strtej = "C:\Users\tejas\OneDrive\Desktop\DevBotAuth4EchoBot_24_1_2020\Demo12_DevBotAuth4EchoBot\Demo12_DevBotAuth4EchoBot\Images\Testimage.png";"
                        card.Body.Add(new AdaptiveContainer()
                        {


                            Items = new List<AdaptiveElement>()
                                {



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
                                                   Color=AdaptiveTextColor.Good
                                              }
                                          }
                                      },
                                      new AdaptiveColumn()
                                      {
                                          Type="Column",
                                          Width="55px",
                                          Items=new List<AdaptiveElement>()
                                          {
                                              new AdaptiveTextBlock()
                                              {
                                                  Type="TextBlock",
                                                  Text="Team",
                                                   Weight=AdaptiveTextWeight.Bolder,
                                                   Color = AdaptiveTextColor.Good

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
                                                   Color =AdaptiveTextColor.Good
                                              }
                                          }
                                      },
                                      new AdaptiveColumn()
                                      {
                                          Type="Column",
                                          Width="130px",
                                          Items=new List<AdaptiveElement>()
                                          {
                                              new AdaptiveTextBlock()
                                              {
                                                  Type="TextBlock",
                                                  Text="Date",
                                                   Weight=AdaptiveTextWeight.Bolder,
                                                   Color =AdaptiveTextColor.Good
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
                                      Width="55px",
                                      Items=items,


                                  },

                                   new AdaptiveColumn()
                                      {
                                          Type="Column",
                                          Width="100px",
                                         Items= items3,
                                      },
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="135px",
                                     Items= items1,
                                  }


                                }


                            },
                            }
                        });


                    }
                    if (reasonintent.Equals("leaveCount") || reasonintent == "leaveCount")
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

                        System.Uri uri = new System.Uri("C:/Users/tejas/OneDrive/Desktop/DevBotAuth4EchoBot_24_1_2020/Demo12_DevBotAuth4EchoBot/Demo12_DevBotAuth4EchoBot/Images/Testimage.png");
                        // string strtej = "C:\Users\tejas\OneDrive\Desktop\DevBotAuth4EchoBot_24_1_2020\Demo12_DevBotAuth4EchoBot\Demo12_DevBotAuth4EchoBot\Images\Testimage.png";"
                        card.Body.Add(new AdaptiveContainer()
                        {


                            Items = new List<AdaptiveElement>()
                                {


                            new AdaptiveTextBlock()
                            {

                                Type="TextBlock",
                               Text="No of Leaves are: "+countOfLeave.ToString(),
                               Height=AdaptiveHeight.Stretch,
                               Color=AdaptiveTextColor.Good,
                               Size=AdaptiveTextSize.Large,
                               Weight=AdaptiveTextWeight.Bolder,
                               HorizontalAlignment=AdaptiveHorizontalAlignment.Left,


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
                                                   Color=AdaptiveTextColor.Good
                                              }
                                          }
                                      },
                                      new AdaptiveColumn()
                                      {
                                          Type="Column",
                                          Width="55px",
                                          Items=new List<AdaptiveElement>()
                                          {
                                              new AdaptiveTextBlock()
                                              {
                                                  Type="TextBlock",
                                                  Text="Team",
                                                   Weight=AdaptiveTextWeight.Bolder,
                                                   Color = AdaptiveTextColor.Good

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
                                                   Color =AdaptiveTextColor.Good
                                              }
                                          }
                                      },
                                      new AdaptiveColumn()
                                      {
                                          Type="Column",
                                          Width="130px",
                                          Items=new List<AdaptiveElement>()
                                          {
                                              new AdaptiveTextBlock()
                                              {
                                                  Type="TextBlock",
                                                  Text="Date",
                                                   Weight=AdaptiveTextWeight.Bolder,
                                                   Color =AdaptiveTextColor.Good
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
                                      Width="55px",
                                      Items=items,


                                  },

                                   new AdaptiveColumn()
                                      {
                                          Type="Column",
                                          Width="100px",
                                         Items= items3,
                                      },
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="135px",
                                     Items= items1,
                                  }


                                }


                            },
                            }
                        });


                    }
                    if (reasonintent.Equals("reason") || reasonintent == "reason")
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
                        for (int i = 0; i < list1.Count; i++)
                        {
                            items1.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].StartDate,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent

                            });
                        }


                        System.Uri uri = new System.Uri("C:/Users/tejas/OneDrive/Desktop/DevBotAuth4EchoBot_24_1_2020/Demo12_DevBotAuth4EchoBot/Demo12_DevBotAuth4EchoBot/Images/Testimage.png");
                        card.Body.Add(new AdaptiveContainer()
                        {

                            Items = new List<AdaptiveElement>()
                    {




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
                                               Color=AdaptiveTextColor.Good

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
                                               Color =AdaptiveTextColor.Good
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
                                               Color=AdaptiveTextColor.Good

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
                                      Width="100px",


                                     Items= items3,
                                  },


                               new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="140px",


                                     Items= items1,
                                  }



                            }


                        },
                    }
                        });


                    }

                    //***************************

                    if (reasonintent.Contains("workingLocations") || reasonintent.Equals("workingLocations") || reasonintent == "workingLocations")
                    {


                        int a = list1.Count;


                        for (int i = 0; i < list1.Count; i++)
                        {

                            dateAdaptive.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].Date,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent
                            });
                        }

                        for (int i = 0; i < list1.Count; i++)
                        {

                            AssignedToAdaptive.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].AssignedTo,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent
                            });
                        }
                        for (int i = 0; i < list1.Count; i++)
                        {

                            PresentAdaptive.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].Present,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent

                            });
                        }
                        for (int i = 0; i < list1.Count; i++)
                        {

                            WorkingLocationAdaptive.Add(new AdaptiveTextBlock()
                            {
                                Text = list1[i].WorkingLocation,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent

                            });
                        }




                        System.Uri uri = new System.Uri("C:/Users/tejas/OneDrive/Desktop/DevBotAuth4EchoBot_24_1_2020/Demo12_DevBotAuth4EchoBot/Demo12_DevBotAuth4EchoBot/Images/Testimage.png");
                        // string strtej = "C:\Users\tejas\OneDrive\Desktop\DevBotAuth4EchoBot_24_1_2020\Demo12_DevBotAuth4EchoBot\Demo12_DevBotAuth4EchoBot\Images\Testimage.png";"
                        card.Body.Add(new AdaptiveContainer()
                        {


                            Items = new List<AdaptiveElement>()
                    {




                          new AdaptiveColumnSet()
                            {
                                Type = "ColumnSet",
                                Height = AdaptiveHeight.Auto,
                                Columns=new List<AdaptiveColumn> ()
                                {
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="120px",

                                      Items=new List<AdaptiveElement>()
                                      {

                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Team Member",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color=AdaptiveTextColor.Good
                                          }
                                      }
                                  }  ,
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="80px",
                                      Items=new List<AdaptiveElement>()
                                      {
                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Present",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color = AdaptiveTextColor.Good

                                          }
                                      }
                                  },
                                    new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="80px",
                                      Items=new List<AdaptiveElement>()
                                      {
                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="Date",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color =AdaptiveTextColor.Good
                                          }
                                      }
                                  },
                                  new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="130px",
                                      Items=new List<AdaptiveElement>()
                                      {
                                          new AdaptiveTextBlock()
                                          {
                                              Type="TextBlock",
                                              Text="WorkingLocation",
                                               Weight=AdaptiveTextWeight.Bolder,
                                               Color =AdaptiveTextColor.Good
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
                                  Width="120px",
                                  Items=AssignedToAdaptive,


                              },

                               new AdaptiveColumn()
                                  {
                                      Type="Column",
                                      Width="80px",
                                     Items= PresentAdaptive,
                                  },

                                 new AdaptiveColumn()
                              {
                                  Type="Column",
                                  Width="80px",
                                 Items= dateAdaptive,
                              }
                               ,
                              new AdaptiveColumn()
                              {
                                  Type="Column",
                                  Width="135px",
                                 Items= WorkingLocationAdaptive,
                              }


                            }


                        },
                    }
                        });


                    }






                    //***********************************

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


                    //Conference room 
                    //*********

                    if (reasonintent == "How")
                    {
                        var findRooms = await graphClient.Me
                      .FindRooms()
                      .Request()
                      .GetAsync();
                        int aa = findRooms.Count;
                        for (int i = 0; i < aa; i++)
                        {
                            String a = findRooms[i].Name;
                            FindRoom.Add(new AdaptiveTextBlock()
                            {
                                Text = a,
                                Size = AdaptiveTextSize.Small,
                                Color = AdaptiveTextColor.Accent
                            });
                        }

                        String ast = "There are " + aa + " Rooms available";

                        card.Body.Add(new AdaptiveContainer()
                        {


                            Items = new List<AdaptiveElement>()
                                {




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
                                                  Text="Conference Rooms",
                                                   Weight=AdaptiveTextWeight.Bolder,
                                                   Color=AdaptiveTextColor.Good
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
                                          Width="500px",

                                          Items=new List<AdaptiveElement>()
                                          {

                                              new AdaptiveTextBlock()
                                              {
                                                  Type="TextBlock",
                                                  Text=ast,
                                                   Weight=AdaptiveTextWeight.Bolder,
                                                   Color=AdaptiveTextColor.Good
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
                                      Width="300px",
                                     Items= FindRoom,
                                  }



                                }


                            },
                            }
                        });

                        var attachment1 = new Microsoft.Bot.Schema.Attachment
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = card,
                        };
                        var reply1 = MessageFactory.Attachment(attachment1);
                        await turnContext.SendActivityAsync(reply1, cancellationToken);

                    }

                    //**************


                    //********************************output on bot for Schedule intent : Adaptive card********************


                    //String roomluisname = "";

                    if (reasonintent == "Schedule")
                    {

                        //******************
                        UAT_LUIS_Entity UAT_LUIS = await GetEntityFromLUIS(turnContext.Activity.Text);
                        List<string> entityList = new List<string>();
                        String HourFromTime = "", HourToTime = "", MinToTime = "", MinFromTime = "", checkspecificdate = "";
                        String FromTime = "", ToTime = "", FromTimeHour = "", FromTimeMin = "", ToTimeHour = "", ToTimeMin = "";
                        for (int i = 0; i < UAT_LUIS.entities.Length; i++)
                        {
                            entityList.Add(UAT_LUIS.entities[i].type.ToString().ToLower());
                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "room")
                            {
                                roomluisname = UAT_LUIS.entities[i].entity.ToString().ToLower();
                            }

                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "date")
                            {
                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                                checkspecificdate = "true"
;
                            }

                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "today")
                            {
                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                            }

                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "the day after tomorrow")
                            {
                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                            }
                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "tomorrow")
                            {
                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                            }




                        }

                        String d = dateluis;


                        if (checkspecificdate == "true")
                        {
                            datevar = Convert.ToDateTime(dateluis);
                        }
                        else if (dateluis == "today")
                        {
                            datevar = DateTime.Today;
                        }
                        else if (dateluis == "tomorrow")
                        {
                            datevar = DateTime.Today.AddDays(1);
                        }
                        else if (dateluis == "the day after tomorrow" || dateluis.Contains("day after tomorrow"))
                        {
                            datevar = DateTime.Today.AddDays(2);
                        }
                        else
                        {
                            datevar = DateTime.Today;
                        }


                        DateTime dateobject;
                        String datestring = "";
                        if (DateTime.TryParse(datevar.ToString(), out dateobject))
                        {

                            datestring = dateobject.ToString("yyyy-MM-ddT");

                        }

                        string datevarparticularstart = datestring + "00" + ":" + "00" + ":" + "01-08:00";
                        string datevarparticularend = datestring + "23" + ":" + "59" + ":" + "00-08:00";
                        var queryparticulardate = new List<QueryOption>()
                        {
                            new QueryOption("startDateTime", datevarparticularstart),
                            new QueryOption("endDateTime", datevarparticularend)
                        };

                        var calendarViewparticulardate = await graphClient.Me.CalendarView
                            .Request(queryparticulardate)
                            .GetAsync();
                        String luiscalenderentity = "";


                        //***********************






                        //                        UAT_LUIS_Entity UAT_LUIS = await GetEntityFromLUIS(turnContext.Activity.Text);
                        //                        List<string> entityList = new List<string>();
                        //                        //String HourFromTime = "", HourToTime = "", MinToTime = "", MinFromTime = "";
                        //                        for (int i = 0; i < UAT_LUIS.entities.Length; i++)
                        //                        {


                        //                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "date")
                        //                            {
                        //                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                        //                            }

                        //                        }







                        //                        if (dateluis != null)
                        //                        {




                        //                            DateTime starttimenew1 = Convert.ToDateTime(dateluis);
                        //                            DateTime dtobj23;
                        //                            String newDateTime111 = "";
                        //                            if (DateTime.TryParse(starttimenew1.ToString(), out dtobj23))
                        //                            {

                        //                                newDateTime111 = dtobj23.ToString("yyyy-MM-ddT");

                        //                            }

                        //                            String try1 = newDateTime111;



                        //                            string datevarparticularstart = try1 + "00" + ":" + "00" + ":" + "01-08:00";
                        //                            string datevarparticularend = try1 + "23" + ":" + "59" + ":" + "00-08:00";

                        //                            var queryparticulardate = new List<QueryOption>()
                        //{
                        //    new QueryOption("startDateTime", datevarparticularstart),
                        //    new QueryOption("endDateTime", datevarparticularend)
                        //};

                        //                            var calendarViewparticulardate = await graphClient.Me.CalendarView
                        //                                .Request(queryparticulardate)
                        //                                .GetAsync();
                        //                            String luiscalenderentity = "";






                        //                            DateTime t = DateTime.Today;
                        //                            string newDateTime = "";
                        //                            DateTime dt;
                        //                            if (DateTime.TryParse(t.ToString(), out dt))
                        //                            {

                        //                                newDateTime = dt.ToString("yyyy-MM-dd");
                        //                            }

                        //                            String starttimeluis = newDateTime;

                        for (int i = 0; i < UAT_LUIS.entities.Length; i++)
                        {
                            entityList.Add(UAT_LUIS.entities[i].type.ToString().ToLower());
                            if (UAT_LUIS.entities[i].entity.ToString().ToLower().Contains("tulip"))
                            {
                                luiscalenderentity = "tulip";

                            }

                            else if (UAT_LUIS.entities[i].entity.ToString().ToLower().Contains("lotus"))
                            {
                                luiscalenderentity = "lotus";

                            }

                            else if (UAT_LUIS.entities[i].entity.ToString().ToLower().Contains("snowdrop"))
                            {
                                luiscalenderentity = "snowdrop";

                            }
                            else if (UAT_LUIS.entities[i].entity.ToString().ToLower().Contains("chanyaky"))
                            {
                                luiscalenderentity = "chanyaky";

                            }
                        }

                        string boolvarcheck = "";

                        if (calendarViewparticulardate.Count!=0)
                        {

                     
                        for (int i = 0; i < calendarViewparticulardate.Count; i++)
                        {


                            if (calendarViewparticulardate[i].Location.DisplayName.ToLower().Contains(luiscalenderentity))
                            {
                                boolvarcheck = "true";

                                string converttostring1 = "";
                                DateTime date1;
                                if (DateTime.TryParse(calendarViewparticulardate[i].Start.DateTime.ToString(), out date1))
                                {

                                    converttostring1 = date1.ToString("HH:mm:ss");
                                }

                                String vartime1 = converttostring1;


                                string converttostring2 = "";
                                DateTime date2;
                                if (DateTime.TryParse(calendarViewparticulardate[i].End.DateTime.ToString(), out date2))
                                {

                                    converttostring2 = date2.ToString("HH:mm:ss");
                                }

                                String vartime2 = converttostring2;


                                EventNameschedule.Add(new AdaptiveTextBlock()
                                {
                                    Text = calendarViewparticulardate[i].Location.DisplayName,
                                    Size = AdaptiveTextSize.Small,
                                    Color = AdaptiveTextColor.Accent
                                });
                                Eventschedulestart.Add(new AdaptiveTextBlock()
                                {
                                    Text = vartime1,
                                    Size = AdaptiveTextSize.Small,
                                    Color = AdaptiveTextColor.Accent
                                });
                                Eventscheduleend.Add(new AdaptiveTextBlock()
                                {
                                    Text = vartime2,
                                    Size = AdaptiveTextSize.Small,
                                    Color = AdaptiveTextColor.Accent
                                });
                                Eventschedulesubject.Add(new AdaptiveTextBlock()
                                {
                                    Text = calendarViewparticulardate[i].Subject,
                                    Size = AdaptiveTextSize.Small,
                                    Color = AdaptiveTextColor.Accent
                                });
                            }
                        }



                          

                            if (boolvarcheck == "true")
                            {
                                card.Body.Add(new AdaptiveContainer()
                                {


                                    Items = new List<AdaptiveElement>()
                                                {




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
                                                                          Text=luiscalenderentity+" Conference Hall",
                                                                           Weight=AdaptiveTextWeight.Bolder,
                                                                           Color=AdaptiveTextColor.Good
                                                                      }
                                                                  }
                                                              }  ,
                                                              new AdaptiveColumn()
                                                              {
                                                                  Type="Column",
                                                                  Width="150px",
                                                                  Items=new List<AdaptiveElement>()
                                                                  {
                                                                      new AdaptiveTextBlock()
                                                                      {
                                                                          Type="TextBlock",
                                                                          Text=DateTime.Today.ToString("dd-MM-yyyy"),
                                                                           Weight=AdaptiveTextWeight.Bolder,
                                                                           Color = AdaptiveTextColor.Good

                                                                      }
                                                                  }
                                                              },

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

                                                                  Items=new List<AdaptiveElement>()
                                                                  {

                                                                      new AdaptiveTextBlock()
                                                                      {
                                                                          Type="TextBlock",
                                                                          Text="Name",
                                                                           Weight=AdaptiveTextWeight.Bolder,
                                                                           Color=AdaptiveTextColor.Good
                                                                      }
                                                                  }
                                                              }  ,
                                                              new AdaptiveColumn()
                                                              {
                                                                  Type="Column",
                                                                  Width="80px",
                                                                  Items=new List<AdaptiveElement>()
                                                                  {
                                                                      new AdaptiveTextBlock()
                                                                      {
                                                                          Type="TextBlock",
                                                                          Text="Start Time",
                                                                           Weight=AdaptiveTextWeight.Bolder,
                                                                           Color = AdaptiveTextColor.Good

                                                                      }
                                                                  }
                                                              },
                                                              new AdaptiveColumn()
                                                              {
                                                                  Type="Column",
                                                                  Width="80px",
                                                                  Items=new List<AdaptiveElement>()
                                                                  {
                                                                      new AdaptiveTextBlock()
                                                                      {
                                                                          Type="TextBlock",
                                                                          Text="End Time",
                                                                           Weight=AdaptiveTextWeight.Bolder,
                                                                           Color = AdaptiveTextColor.Good

                                                                      }
                                                                  }
                                                              },

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
                                                              Items=Eventschedulesubject,


                                                          },

                                                           new AdaptiveColumn()
                                                              {
                                                                  Type="Column",
                                                                  Width="80px",
                                                                 Items= Eventschedulestart,
                                                              },

                                                             new AdaptiveColumn()
                                                          {
                                                              Type="Column",
                                                              Width="80px",
                                                             Items= Eventscheduleend,
                                                          }



                                                        }


                                                    },
                                                }
                                });
                            }

                            else
                            {
                                card.Body.Add(new AdaptiveContainer()
                                {


                                    Items = new List<AdaptiveElement>()
                                                {




                                                      new AdaptiveColumnSet()
                                                        {
                                                            Type = "ColumnSet",
                                                            Height = AdaptiveHeight.Auto,
                                                            Columns=new List<AdaptiveColumn> ()
                                                            {
                                                              new AdaptiveColumn()
                                                              {
                                                                  Type="Column",
                                                                  Width="500px",

                                                                  Items=new List<AdaptiveElement>()
                                                                  {

                                                                      new AdaptiveTextBlock()
                                                                      {
                                                                          Type="TextBlock",
                                                                          Text=" No scheduling for "+luiscalenderentity+" hall",
                                                                           Weight=AdaptiveTextWeight.Bolder,
                                                                           Color=AdaptiveTextColor.Good
                                                                      }
                                                                  }
                                                              }


                                                             }
                                                      },








                                                }
                                });
                            }


                        }

                        else
                        {
                            card.Body.Add(new AdaptiveContainer()
                            {


                                Items = new List<AdaptiveElement>()
                                                {




                                                      new AdaptiveColumnSet()
                                                        {
                                                            Type = "ColumnSet",
                                                            Height = AdaptiveHeight.Auto,
                                                            Columns=new List<AdaptiveColumn> ()
                                                            {
                                                              new AdaptiveColumn()
                                                              {
                                                                  Type="Column",
                                                                  Width="500px",

                                                                  Items=new List<AdaptiveElement>()
                                                                  {

                                                                      new AdaptiveTextBlock()
                                                                      {
                                                                          Type="TextBlock",
                                                                          Text=" No scheduling for "+luiscalenderentity+" hall",
                                                                           Weight=AdaptiveTextWeight.Bolder,
                                                                           Color=AdaptiveTextColor.Good
                                                                      }
                                                                  }
                                                              }


                                                             }
                                                      },








                                                }
                            });
                        }

                        var attachment1 = new Microsoft.Bot.Schema.Attachment
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = card,
                        };
                        var reply1 = MessageFactory.Attachment(attachment1);
                        await turnContext.SendActivityAsync(reply1, cancellationToken);

                        //                            }

                        //                            else
                        //                            {
                        //                                card.Body.Add(new AdaptiveContainer()
                        //                                {


                        //                                    Items = new List<AdaptiveElement>()
                        //                    {




                        //                          new AdaptiveColumnSet()
                        //                            {
                        //                                Type = "ColumnSet",
                        //                                Height = AdaptiveHeight.Auto,
                        //                                Columns=new List<AdaptiveColumn> ()
                        //                                {
                        //                                  new AdaptiveColumn()
                        //                                  {
                        //                                      Type="Column",
                        //                                      Width="500px",

                        //                                      Items=new List<AdaptiveElement>()
                        //                                      {

                        //                                          new AdaptiveTextBlock()
                        //                                          {
                        //                                              Type="TextBlock",
                        //                                              Text=" No scheduling for "+luiscalenderentity+" hall",
                        //                                               Weight=AdaptiveTextWeight.Bolder,
                        //                                               Color=AdaptiveTextColor.Good
                        //                                          }
                        //                                      }
                        //                                  }


                        //                                 }
                        //                          },








                        //                    }
                        //                                });
                        //                            }

                        //                            var attachment1 = new Microsoft.Bot.Schema.Attachment
                        //                            {
                        //                                ContentType = AdaptiveCard.ContentType,
                        //                                Content = card,
                        //                            };
                        //                            var reply1 = MessageFactory.Attachment(attachment1);
                        //                            await turnContext.SendActivityAsync(reply1, cancellationToken);
                        //                        }

                        //                        else
                        //                        {
                        //                            String luiscalenderentity = "";


                        //                            var queryOptions = new List<QueryOption>()
                        //            {
                        //                new QueryOption("startDateTime", "2020-02-20T00:00:01-08:00"),
                        //                new QueryOption("endDateTime", "2020-02-20T23:59:59-08:00")
                        //            };

                        //                            var calendarView = await graphClient.Me.Calendar.CalendarView
                        //                                .Request(queryOptions)
                        //                                .GetAsync();




                        //                            DateTime t = DateTime.Today;
                        //                            string newDateTime = "";
                        //                            DateTime dt;
                        //                            if (DateTime.TryParse(t.ToString(), out dt))
                        //                            {

                        //                                newDateTime = dt.ToString("yyyy-MM-dd");
                        //                            }

                        //                            String starttimeluis = newDateTime;

                        //                            for (int i = 0; i < UAT_LUIS.entities.Length; i++)
                        //                            {
                        //                                entityList.Add(UAT_LUIS.entities[i].type.ToString().ToLower());
                        //                                if (UAT_LUIS.entities[i].entity.ToString().ToLower().Contains("tulip"))
                        //                                {
                        //                                    luiscalenderentity = "tulip";

                        //                                }

                        //                                else if (UAT_LUIS.entities[i].entity.ToString().ToLower().Contains("lotus"))
                        //                                {
                        //                                    luiscalenderentity = "lotus";

                        //                                }

                        //                                else if (UAT_LUIS.entities[i].entity.ToString().ToLower().Contains("snowdrop"))
                        //                                {
                        //                                    luiscalenderentity = "snowdrop";

                        //                                }
                        //                                else if (UAT_LUIS.entities[i].entity.ToString().ToLower().Contains("chanyaky"))
                        //                                {
                        //                                    luiscalenderentity = "chanyaky";

                        //                                }
                        //                            }


                        //                            string boolvarcheck = "";
                        //                            for (int i = 0; i < calendarView.Count; i++)
                        //                            {


                        //                                if (calendarView[i].Location.DisplayName.ToLower().Contains(luiscalenderentity))
                        //                                {
                        //                                    boolvarcheck = "true";

                        //                                    string converttostring1 = "";
                        //                                    DateTime date1;
                        //                                    if (DateTime.TryParse(calendarView[i].Start.DateTime.ToString(), out date1))
                        //                                    {

                        //                                        converttostring1 = date1.ToString("HH:mm:ss");
                        //                                    }

                        //                                    String vartime1 = converttostring1;


                        //                                    string converttostring2 = "";
                        //                                    DateTime date2;
                        //                                    if (DateTime.TryParse(calendarView[i].End.DateTime.ToString(), out date2))
                        //                                    {

                        //                                        converttostring2 = date2.ToString("HH:mm:ss");
                        //                                    }

                        //                                    String vartime2 = converttostring2;


                        //                                    EventNameschedule.Add(new AdaptiveTextBlock()
                        //                                    {
                        //                                        Text = calendarView[i].Location.DisplayName,
                        //                                        Size = AdaptiveTextSize.Small,
                        //                                        Color = AdaptiveTextColor.Accent
                        //                                    });
                        //                                    Eventschedulestart.Add(new AdaptiveTextBlock()
                        //                                    {
                        //                                        Text = vartime1,
                        //                                        Size = AdaptiveTextSize.Small,
                        //                                        Color = AdaptiveTextColor.Accent
                        //                                    });
                        //                                    Eventscheduleend.Add(new AdaptiveTextBlock()
                        //                                    {
                        //                                        Text = vartime2,
                        //                                        Size = AdaptiveTextSize.Small,
                        //                                        Color = AdaptiveTextColor.Accent
                        //                                    });
                        //                                    Eventschedulesubject.Add(new AdaptiveTextBlock()
                        //                                    {
                        //                                        Text = calendarView[i].Subject,
                        //                                        Size = AdaptiveTextSize.Small,
                        //                                        Color = AdaptiveTextColor.Accent
                        //                                    });
                        //                                }
                        //                            }

                        //                            if (boolvarcheck == "true")
                        //                            {
                        //                                card.Body.Add(new AdaptiveContainer()
                        //                                {


                        //                                    Items = new List<AdaptiveElement>()
                        //                    {




                        //                          new AdaptiveColumnSet()
                        //                            {
                        //                                Type = "ColumnSet",
                        //                                Height = AdaptiveHeight.Auto,
                        //                                Columns=new List<AdaptiveColumn> ()
                        //                                {
                        //                                  new AdaptiveColumn()
                        //                                  {
                        //                                      Type="Column",
                        //                                      Width="200px",

                        //                                      Items=new List<AdaptiveElement>()
                        //                                      {

                        //                                          new AdaptiveTextBlock()
                        //                                          {
                        //                                              Type="TextBlock",
                        //                                              Text=luiscalenderentity+" Conference Hall",
                        //                                               Weight=AdaptiveTextWeight.Bolder,
                        //                                               Color=AdaptiveTextColor.Good
                        //                                          }
                        //                                      }
                        //                                  }  ,
                        //                                  new AdaptiveColumn()
                        //                                  {
                        //                                      Type="Column",
                        //                                      Width="150px",
                        //                                      Items=new List<AdaptiveElement>()
                        //                                      {
                        //                                          new AdaptiveTextBlock()
                        //                                          {
                        //                                              Type="TextBlock",
                        //                                              Text=DateTime.Today.ToString("dd-MM-yyyy"),
                        //                                               Weight=AdaptiveTextWeight.Bolder,
                        //                                               Color = AdaptiveTextColor.Good

                        //                                          }
                        //                                      }
                        //                                  },

                        //                                 }
                        //                          },

                        //                          new AdaptiveColumnSet()
                        //                            {
                        //                                Type = "ColumnSet",
                        //                                Height = AdaptiveHeight.Auto,
                        //                                Columns=new List<AdaptiveColumn> ()
                        //                                {
                        //                                  new AdaptiveColumn()
                        //                                  {
                        //                                      Type="Column",
                        //                                      Width="200px",

                        //                                      Items=new List<AdaptiveElement>()
                        //                                      {

                        //                                          new AdaptiveTextBlock()
                        //                                          {
                        //                                              Type="TextBlock",
                        //                                              Text="Name",
                        //                                               Weight=AdaptiveTextWeight.Bolder,
                        //                                               Color=AdaptiveTextColor.Good
                        //                                          }
                        //                                      }
                        //                                  }  ,
                        //                                  new AdaptiveColumn()
                        //                                  {
                        //                                      Type="Column",
                        //                                      Width="80px",
                        //                                      Items=new List<AdaptiveElement>()
                        //                                      {
                        //                                          new AdaptiveTextBlock()
                        //                                          {
                        //                                              Type="TextBlock",
                        //                                              Text="Start Time",
                        //                                               Weight=AdaptiveTextWeight.Bolder,
                        //                                               Color = AdaptiveTextColor.Good

                        //                                          }
                        //                                      }
                        //                                  },
                        //                                  new AdaptiveColumn()
                        //                                  {
                        //                                      Type="Column",
                        //                                      Width="80px",
                        //                                      Items=new List<AdaptiveElement>()
                        //                                      {
                        //                                          new AdaptiveTextBlock()
                        //                                          {
                        //                                              Type="TextBlock",
                        //                                              Text="End Time",
                        //                                               Weight=AdaptiveTextWeight.Bolder,
                        //                                               Color = AdaptiveTextColor.Good

                        //                                          }
                        //                                      }
                        //                                  },

                        //                                 }
                        //                          },




                        //                        new AdaptiveColumnSet()
                        //                        {
                        //                            Type = "ColumnSet",
                        //                            Height = AdaptiveHeight.Auto,

                        //                            Columns=new List<AdaptiveColumn> ()
                        //                            {


                        //                              new AdaptiveColumn()
                        //                              {
                        //                                  Type="Column",
                        //                                  Width="200px",
                        //                                  Items=Eventschedulesubject,


                        //                              },

                        //                               new AdaptiveColumn()
                        //                                  {
                        //                                      Type="Column",
                        //                                      Width="80px",
                        //                                     Items= Eventschedulestart,
                        //                                  },

                        //                                 new AdaptiveColumn()
                        //                              {
                        //                                  Type="Column",
                        //                                  Width="80px",
                        //                                 Items= Eventscheduleend,
                        //                              }



                        //                            }


                        //                        },
                        //                    }
                        //                                });
                        //                            }

                        //                            else
                        //                            {
                        //                                card.Body.Add(new AdaptiveContainer()
                        //                                {


                        //                                    Items = new List<AdaptiveElement>()
                        //                    {




                        //                          new AdaptiveColumnSet()
                        //                            {
                        //                                Type = "ColumnSet",
                        //                                Height = AdaptiveHeight.Auto,
                        //                                Columns=new List<AdaptiveColumn> ()
                        //                                {
                        //                                  new AdaptiveColumn()
                        //                                  {
                        //                                      Type="Column",
                        //                                      Width="500px",

                        //                                      Items=new List<AdaptiveElement>()
                        //                                      {

                        //                                          new AdaptiveTextBlock()
                        //                                          {
                        //                                              Type="TextBlock",
                        //                                              Text=" No scheduling for "+luiscalenderentity+" hall",
                        //                                               Weight=AdaptiveTextWeight.Bolder,
                        //                                               Color=AdaptiveTextColor.Good
                        //                                          }
                        //                                      }
                        //                                  }


                        //                                 }
                        //                          },








                        //                    }
                        //                                });
                        //                            }

                        //                            var attachment1 = new Microsoft.Bot.Schema.Attachment
                        //                            {
                        //                                ContentType = AdaptiveCard.ContentType,
                        //                                Content = card,
                        //                            };
                        //                            var reply1 = MessageFactory.Attachment(attachment1);
                        //                            await turnContext.SendActivityAsync(reply1, cancellationToken);
                        //                        }

                    }





                    //*******


                    if (reasonintent == "Available")

                    {
                       

                        

                        UAT_LUIS_Entity UAT_LUIS = await GetEntityFromLUIS(turnContext.Activity.Text);
                        List<string> entityList = new List<string>();
                        String HourFromTime = "", HourToTime = "", MinToTime = "", MinFromTime = "",checkspecificdate="";
                        String FromTime="", ToTime="", FromTimeHour="", FromTimeMin="", ToTimeHour="", ToTimeMin="";
                        for (int i = 0; i < UAT_LUIS.entities.Length; i++)
                        {
                            entityList.Add(UAT_LUIS.entities[i].type.ToString().ToLower());
                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "room")
                            {
                                roomluisname = UAT_LUIS.entities[i].entity.ToString().ToLower();
                            }

                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "date")
                            {
                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                                checkspecificdate = "true"
;                            }

                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "today")
                            {
                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                            }

                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "the day after tomorrow")
                            {
                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                            }
                            if (UAT_LUIS.entities[i].type.ToString().ToLower() == "tomorrow")
                            {
                                dateluis = UAT_LUIS.entities[i].entity.ToString().ToLower();
                            }




                        }

                        //*****

                        
                        //************code begin here******************

                        String d = dateluis;


                        if (checkspecificdate == "true")
                        {
                            datevar = Convert.ToDateTime(dateluis);
                        }
                        else if (dateluis == "today")
                        {
                            datevar = DateTime.Today;
                        }
                        else if (dateluis == "tomorrow")
                        {
                            datevar = DateTime.Today.AddDays(1);
                        }
                        else if (dateluis == "the day after tomorrow")
                        {
                            datevar = DateTime.Today.AddDays(2);
                        }
                        else
                        {
                            datevar = DateTime.Today;
                        }

                        ///**********convertion
                        DateTime dateobject;
                        String datestring = "";
                        if (DateTime.TryParse(datevar.ToString(), out dateobject))
                        {

                            datestring = dateobject.ToString("yyyy-MM-ddT");

                        }


                        //**********************


                        String teju1="", teju2="", temp="";

                        if (entityList.Contains("fromtime") == true)
                        {

                            var indexVal = Enumerable.Range(0, entityList.Count)
                                         .Where(i => entityList[i] == "fromtime")
                                         .ToList();
                            if (indexVal.Count == 1)
                            {
                                teju1 = UAT_LUIS.entities[indexVal[0]].entity.ToString();


                                if(teju1.Contains("pm"))
                                {

                                    
                                        int index2 = teju1.IndexOf("pm");
                                        if (index2 != -1)
                                        {
                                            ResultFromTime = teju1.Remove(index2);


                                        }
                                    temp = ResultFromTime;

                                }

                               

                                
                                else
                                {
                                int index2 = teju1.IndexOf("am");
                                if (index2 != -1)
                                {
                                    ResultFromTime = teju1.Remove(index2);


                                }
                                temp = ResultFromTime;

                            }

                           


                        






                                if(teju1.Contains("am"))
                                {
                                    teju2 = "00am";
                                }
                                else
                                {
                                    teju2 = "00pm";
                                }
                                
                                teju1 = temp;
                 
                            }
                            else
                            {
                                for (int i = 0; i < indexVal.Count; i++)
                                {
                                    if (UAT_LUIS.entities[indexVal[i]].role.ToString().Equals("HourTime"))
                                    {
                                        teju1 = UAT_LUIS.entities[indexVal[i]].entity.ToString();



                                    }
                                    else if (UAT_LUIS.entities[indexVal[i]].role.ToString().Equals("MinTime"))
                                        teju2 = UAT_LUIS.entities[indexVal[i]].entity.ToString();

                                         

                                }
                            }


                        }





                        if (teju2.Contains("pm"))
                        {
                            int index2 = teju2.IndexOf("pm");
                            if (index2 != -1)
                            {
                                ResultFromTime = teju2.Remove(index2) ;
                              
                           
                            }

                                                                                 
                        }
                        if (teju2.Contains("am"))
                        {
                            int index2 = teju2.IndexOf("am");
                            if (index2 != -1)
                            {
                                ResultFromTime = teju2.Remove(index2);


                            }


                        }

                        String asd;
                        if (teju2.Contains("pm"))
                        {

                            asd = teju1 + ":"+ResultFromTime+" PM";
                        }
                        else
                        {
                             asd = teju1 +":"+ ResultFromTime+" AM";
                        }

                      ;
                        DateTime dateobject11;
                        String datestring11 = "";
                        if (DateTime.TryParse(asd.ToString(), out dateobject11))
                        {

                            datestring11 = dateobject11.ToString("HH:mm");

                        }



                        //**************


                        String teju21 = "", teju22 = "";

                        if (entityList.Contains("totime") == true)
                        {

                            var indexVal = Enumerable.Range(0, entityList.Count)
                                         .Where(i => entityList[i] == "totime")
                                         .ToList();
                            if (indexVal.Count == 1)
                            {
                                teju21 = UAT_LUIS.entities[indexVal[0]].entity.ToString();


                                if (teju21.Contains("pm"))
                                {


                                    int index2 = teju21.IndexOf("pm");
                                    if (index2 != -1)
                                    {
                                        ResultFromTime = teju21.Remove(index2);


                                    }
                                    temp = ResultFromTime;

                                }




                                else
                                {
                                    int index2 = teju21.IndexOf("am");
                                    if (index2 != -1)
                                    {
                                        ResultFromTime = teju21.Remove(index2);


                                    }
                                    temp = ResultFromTime;

                                }







                                String g = teju21;

                                if (teju21.Contains("am"))
                                {
                                    teju22 = "00am";
                                }

                                else
                                {
                                    teju22 = "00pm";
                                }
                               
                                teju21 = temp;

                            }
                            else
                            {
                                for (int i = 0; i < indexVal.Count; i++)
                                {
                                    if (UAT_LUIS.entities[indexVal[i]].role.ToString().Equals("HourTime1"))
                                    {
                                        teju21 = UAT_LUIS.entities[indexVal[i]].entity.ToString();



                                    }
                                    else if (UAT_LUIS.entities[indexVal[i]].role.ToString().Equals("MinTime1"))
                                        teju22 = UAT_LUIS.entities[indexVal[i]].entity.ToString();



                                }
                            }


                        }



                        if (teju22.Contains("pm"))
                        {
                            int index2 = teju22.IndexOf("pm");
                            if (index2 != -1)
                            {
                                ResultFromTime = teju22.Remove(index2);


                            }


                        }
                        if (teju22.Contains("am"))
                        {
                            int index2 = teju22.IndexOf("am");
                            if (index2 != -1)
                            {
                                ResultFromTime = teju22.Remove(index2);


                            }


                        }

                        String asd1;
                        if (teju22.Contains("pm"))
                        {

                            asd1 = teju21 + ":" + ResultFromTime + " PM";
                        }
                        else
                        {
                            asd1 = teju21 + ":" + ResultFromTime + " AM";
                        }






                        DateTime dateobject111;
                        String datestring111 = "";
                        if (DateTime.TryParse(asd1.ToString(), out dateobject111))
                        {

                            datestring111 = dateobject111.ToString("HH:mm");

                        }

                        string datestart = datestring + datestring11 + ":" + "00-08:00";
                        string dateend = datestring + datestring111 + ":" + "00-08:00";



                        var querydate = new List<QueryOption>()
{
    new QueryOption("startDateTime", datestart),
    new QueryOption("endDateTime", dateend)
};

                        var calendarViewdate = await graphClient.Me.CalendarView
                            .Request(querydate)
                            .GetAsync();



                        if (calendarViewdate.Count != 0)
                        {
                            dateresultoverlap = Convert.ToDateTime(calendarViewdate[0].End.DateTime).AddHours(-8);
                        }


                        string s1 = Convert.ToString(dateresultoverlap);
                        DateTime dtobj1;
                        if (DateTime.TryParse(s1, out dtobj1))
                        {

                            newDateTime1 = dtobj1.ToString("yyyy-MM-ddTHH:mm");

                        }

                        string date1 = newDateTime1;

                        string date2 = datestring + new1 + ":" + new2;



                        if (calendarViewdate.Count != 0)
                        {

                         

                            if (date1 == date2)
                            {
                                botanswer = roomluisname + " is  available";
                            }
                            else
                            {

                                for (int i = 0; i < calendarViewdate.Count; i++)
                                {
                                    String a = roomluisname.ToLower();
                                    String b = calendarViewdate[i].Location.DisplayName.ToLower();

                                    if (roomluisname.ToLower().Contains("tulip") == true && calendarViewdate[i].Location.DisplayName.ToLower().Contains("tulip") == true)
                                    {
                                        botanswer = roomluisname + " is not available";
                                    }
                                    else if (calendarViewdate[i].Location.DisplayName.ToLower().Contains("lotus") == true && roomluisname.ToLower().Contains("lotus") == true)
                                    {
                                        botanswer = roomluisname + " is not available";
                                    }
                                    else if (calendarViewdate[i].Location.DisplayName.ToLower().Contains("snowdrop") == true && roomluisname.ToLower().Contains("snowdrop") == true)
                                    {
                                        botanswer = roomluisname + " is not available";
                                    }
                                    else if (calendarViewdate[i].Location.DisplayName.ToLower().Contains("chanyaky") == true && roomluisname.ToLower().Contains("chanyaky") == true)
                                    {
                                        botanswer = roomluisname + " is not available";
                                    }
                                    else
                                    {
                                        botanswer = roomluisname + " is  available";

                                    }

                                }

                            }

                        }
                        else
                        {


                            botanswer = roomluisname + " is   available";

                        }





                                   






                    }
                            

                    break;
            }


            await turnContext.SendActivityAsync(CreateActivityWithTextAndSpeak($" " + botanswer), cancellationToken);

        }

        //LUIS
        private static async Task<List<Leave>> LuisDatabaSourceCalls(string Query)
        {
            string sysFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            string UserAns = "";
            UAT_LUIS_Entity UAT_LUIS = await GetEntityFromLUIS(Query);
            List<Leave> list1 = new List<Leave>();



            if (!UAT_LUIS.intents.Equals(null))
            {
                if (UAT_LUIS.intents.Length > 0 && UAT_LUIS.entities.Length > 0)
                {
                    string priorEntity = "";

                    priorEntity = UAT_LUIS.entities[0].type.ToString();
                    intentEntity = "Intent: " + UAT_LUIS.intents[0].intent.ToString() + "  Entity: " + priorEntity;
                    reasonintent = UAT_LUIS.intents[0].intent.ToString();

                    list1 = await GetDynamicQueryForSP(UAT_LUIS);


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

                string LuisAppID = "4ddf4dea-2a2d-43f1-8c7c-bd57e5fec74e";// "9945e320 -dd96-4e24-a8f8-482d06756acd";//"d5e2fec8-e3c3-4d3d-bcc2-eaea8530e8a2"; //"11838aa9-ca07-4c63-9206-7e422f3a50bd";/* "16ac5ae9-a86f-4210-93fe-010b387e1262";*///"ab167f38-5367-421f-8e2c-cfa91d36ca19";//// "11838aa9-ca07-4c63-9206-7e422f3a50bd";//"3d87a0d1-a41c-4c2c-b48b-a9bd8c308e61";//"5e376704-da8b-406e-9c68-a7ec15ea29eb";//"6473c60c-64d4-4b4e-895c-946deb9f27bd";//"0e4f70af-c3af-414b-a28f-e9861b17d3ab";//"6473c60c-64d4-4b4e-895c-946deb9f27bd";//"6f963293-9c7a-4393-87eb-9d9a50459662";//ConfigurationManager.AppSettings["aad:LuisAppId"];//["aad:LuisAppId"];
                string LuisSecretKey = "0c856eabd50c4dd0ba0078c08d0ee158";//"784fb0f86ac24d6cb3b75e74191419e5";
//";//"619b66fbeed048e586e81793f5a69925"; /*"619b66fbeed048e586e81793f5a69925";*///"e7301392690a40e387f2a300899d9a4d";////"05b8b88ba1b7480388969d921f5d908c";//"e7301392690a40e387f2a300899d9a4d";//"9f0672f0109247928a0144ca4a910219";//"619b66fbeed048e586e81793f5a69925";//ConfigurationManager.AppSettings["aad:LuisSecretKey"];
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
            string manager = "", PName = "", query = "", labelToMakeQuery = "dateRange", workingLocation = "";
            string userAns = "";
            int noOfLocation = 0;
            SpListName.Add("LeaveManagement");
            SpListName.Add("Attendance");

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


                int.TryParse(firstItem, out noOfDuration);


            }
            if (entityList.Contains("date") == true)
            {

                var indexVal = Enumerable.Range(0, entityList.Count)
                             .Where(i => entityList[i] == "date")
                             .ToList();
                if (indexVal.Count == 1)
                {
                    leaveEndDate = leaveStartDate = Convert.ToDateTime(UAT_LUIS.entities[indexVal[0]].entity.ToString().Replace(" ", string.Empty)).ToString("yyyy-MM-ddTHH:mm:ssZ");

                }
                else
                {
                    for (int i = 0; i < indexVal.Count; i++)
                    {
                        if (UAT_LUIS.entities[indexVal[i]].role.ToString().Equals("FromDate"))
                        {
                            leaveStartDate = Convert.ToDateTime(UAT_LUIS.entities[indexVal[i]].entity.ToString().Replace(" ", string.Empty)).ToString("yyyy-MM-ddTHH:mm:ssZ");

                            // userInput = "start:" + UAT_LUIS.entities[indexVal[i]].entity.ToString() + " End:" + leaveStartDate+"Day: "+ Convert.ToDateTime(leaveStartDate).Day.ToString()+" Month: "+ Convert.ToDateTime(leaveStartDate).Month.ToString();
                        }
                        else if (UAT_LUIS.entities[i].role.ToString().Equals("ToDate"))
                            leaveEndDate = Convert.ToDateTime(UAT_LUIS.entities[indexVal[i]].entity.ToString().Replace(" ", string.Empty)).ToString("yyyy-MM-ddTHH:mm:ssZ");

                    }
                }


            }
            if (entityList.Contains("clientlocation") == true)
            {
                workingLocation = "Client Location";
                noOfLocation++;
            }
            if (entityList.Contains("office") == true)
            {
                if (workingLocation == "")
                    workingLocation = "Office";
                else
                    workingLocation = workingLocation + "_" + "Office";
                noOfLocation++;



            }
            if (entityList.Contains("wfh") == true)
            {
                if (workingLocation == "")
                    workingLocation = "Home";
                else
                    workingLocation = workingLocation + "_" + "Home";
                noOfLocation++;
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
                        if (noOfDuration != 0)
                        {
                            leaveStartDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddDays(-7 * noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        else
                        {
                            leaveStartDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
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
                        else


                        {
                            leaveStartDate = DateTime.UtcNow.AddDays(-(int)DateTime.UtcNow.DayOfWeek).AddDays(7).AddSeconds(+1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = DateTime.UtcNow.AddDays(+(int)DateTime.UtcNow.DayOfWeek).AddDays(6).ToString("yyyy-MM-ddTHH:mm:ssZ");
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
                            leaveStartDate = DateTime.UtcNow.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ssZ");
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
                            //leaveEndDate = DateTime.UtcNow.AddMonths(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(-1).Year, DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(-1).Month, 1).AddMonths(1).AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");

                            labelToMakeQuery = "dateRange";
                        }
                        // leaveEndDate = new DateTime(DateTime.UtcNow.AddYears(noOfDuration).Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        else
                        {
                            leaveStartDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            // leaveEndDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(-1).Year, DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(-1).Month, 1).AddMonths(1).AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");

                            // leaveEndDate = new DateTime(DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).Year, DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).Month, 1).AddMonths(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");

                            labelToMakeQuery = "dateRange";
                        }
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
                            leaveStartDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            //leaveEndDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            //leaveEndDate =new DateTime (DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).Year,31, DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).Month).AddMonths(noOfDuration).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(noOfDuration).Year, DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(noOfDuration).Month, 1).AddMonths(noOfDuration - 1).AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            //new DateTime(origDT.Year, origDT.Month, 1).AddMonths(1).AddDays(-1);
                            labelToMakeQuery = "dateRange";
                        }
                        //leaveEndDate = new DateTime(DateTime.UtcNow.AddYears(noOfDuration).Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        else
                        {
                            leaveStartDate = DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(1).Year, DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).AddMonths(1).Month, 1).AddMonths(1).AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            // leaveEndDate = new DateTime(DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).Year,31, DateTime.UtcNow.AddDays(1 - DateTime.UtcNow.Day).Month, 1).AddMonths(1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
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
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddYears(-1).Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        else
                        {
                            leaveStartDate = new DateTime(DateTime.UtcNow.AddYears(-1).Year, 1, 1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddYears(-1).Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        break;
                    }
                case "next year":
                    {
                        if (noOfDuration != 0)
                        {
                            // case "next year":
                            leaveStartDate = new DateTime(DateTime.UtcNow.AddYears(1).Year, 1, 1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddYears(noOfDuration).Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        else
                        {
                            leaveStartDate = new DateTime(DateTime.UtcNow.AddYears(1).Year, 1, 1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            leaveEndDate = new DateTime(DateTime.UtcNow.AddYears(1).Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                            labelToMakeQuery = "dateRange";
                        }
                        break;
                    }
                case "current year":
                    { //  case "current year":
                        leaveStartDate = new DateTime(DateTime.UtcNow.Year, 1, 1).ToString("yyyy-MM-ddTHH:mm:ssZ");//DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                        leaveEndDate = new DateTime(DateTime.UtcNow.Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                        labelToMakeQuery = "dateRange";

                        break;
                    }
                // case "next to next":
                // {
                // leaveStartDate = new DateTime(DateTime.UtcNow.Year, 1, 1).ToString("yyyy-MM-ddTHH:mm:ssZ");//DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                //  leaveEndDate = new DateTime(DateTime.UtcNow.Year, 12, 31).ToString("yyyy-MM-ddTHH:mm:ssZ");
                // labelToMakeQuery = "dateRange";
                // }

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
                if (UAT_LUIS.intents[0].intent.ToString().Equals("leave"))
                    labelToMakeQuery = "dateRangeWithPName";
                else if (UAT_LUIS.intents[0].intent.ToString().Equals("reason"))
                    labelToMakeQuery = "reasonForPName";
                else if (UAT_LUIS.intents[0].intent.ToString().Equals("leaveCount"))
                    labelToMakeQuery = "leaveCountForPName";
            }
            //}
            switch (UAT_LUIS.intents[0].intent.ToString())
            {
                case "leave":
                    switch (labelToMakeQuery)
                    {
                        case "dateRange":
                            {

                                query = string.Format("<View><Query><ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='Title' /></ViewFields><Where><And>" +
                                 "<Leq><FieldRef Name='StartDate' /><Value Type='DateTime'>{0}</Value></Leq>" +
                                 "<Or><Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{0}</Value></Geq>" +
                                 "<Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{1}</Value></Geq>" +
                                 "</Or></And></Where></Query></View>", leaveEndDate, leaveStartDate);

                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[0]);

                                break;
                            }
                        case "dateRangeWithManager":
                            {
                                query = string.Format("<View><Query><ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='Title' /></ViewFields><Where><And>" +
                                 "<And><Leq><FieldRef Name='StartDate' /><Value Type='DateTime'>{0}</Value></Leq>" +
                                 "<Contains><FieldRef Name='Manager' /><Value Type='UserMulti'>{2}</Value></Contains>" +
                                 "</And><Or><Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{0}</Value></Geq>" +
                                 "<Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{1}</Value></Geq>" +
                                 "</Or></And></Where></Query></View>", leaveEndDate, leaveStartDate, manager);
                                //TM  userAns = await GetAnsDetailsFromSharepoint(query, SpListName[0]);

                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[0]);
                                if (userAns == "")
                                    userAns = "No data found for specific date";
                                break;
                            }
                        //check
                        case "dateRangeWithPName":
                            {
                                query = string.Format("<View><Query><ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='Title' /></ViewFields><Where><And>" +
                                 "<And><Leq><FieldRef Name='StartDate' /><Value Type='DateTime'>{0}</Value></Leq>" +
                                 "<Contains><FieldRef Name='Author' /><Value Type='UserMulti'>{2}</Value></Contains>" +
                                 "</And><Or><Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{0}</Value></Geq>" +
                                 "<Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{1}</Value></Geq>" +
                                 "</Or></And></Where></Query></View>", leaveEndDate, leaveStartDate, PName);
                                //TM  userAns = await GetAnsDetailsFromSharepoint(query, SpListName[0]);

                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[0]);
                                if (userAns == "")
                                    userAns = "No data found for specific date";
                                break;
                            }
                    }
                    break;

                case "reason":
                    switch (labelToMakeQuery)
                    {
                        case "reasonForPName":
                            {
                                //    query = string.Format("<View><Query><ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='Title' /></ViewFields><Where><And>" +
                                //"<Geq><FieldRef Name='StartDate' /><Value Type='DateTime'>2020-02-06T18:30:00Z</Value></Geq>" +
                                //"<Leq><FieldRef Name='EndDate' /><Value Type='DateTime'>{1}</Value></Leq>" +
                                //"</And></Where></Query></View>", leaveStartDate, leaveEndDate);
                                query = string.Format("<View><Query><ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='Title' /></ViewFields><Where><And>" +
                                 "<And><Leq><FieldRef Name='StartDate' /><Value Type='DateTime'>{0}</Value></Leq>" +
                                 "<Contains><FieldRef Name='Author' /><Value Type='User'>{2}</Value></Contains>" +
                                 "</And><Or><Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{0}</Value></Geq>" +
                                 "<Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{1}</Value></Geq>" +
                                 "</Or></And></Where></Query></View>", leaveEndDate, leaveStartDate, PName);

                                // userAns = await GetAnsDetailsFromSharepoint(query, SpListName[1]);
                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[0]);

                                //if (userAns == "")
                                //    userAns = "No data found for specific date";
                                break;
                            }


                    }
                    break;
                //leaveCount
                case "leaveCount":
                    switch (labelToMakeQuery)
                    {
                        case "leaveCountForPName":
                            {
                                if (leaveEndDate.Equals("") && leaveStartDate.Equals(""))
                                {
                                    query = string.Format("<View><Query><Where><Contains><FieldRef Name='Author' /><Value Type='User'>{0}</Value></Contains></Where></Query>" +
                                        "<ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='TestColumn' /></ViewFields>" +
                                        "<QueryOptions /></View>", PName);

                                }
                                else
                                {
                                    query = string.Format("<View><Query><ViewFields><FieldRef Name='LeaveType' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Manager' /><FieldRef Name='RequestStatus' /><FieldRef Name='Author' /><FieldRef Name='Team' /><FieldRef Name='Title' /></ViewFields><Where><And>" +
                                       "<And><Leq><FieldRef Name='StartDate' /><Value Type='DateTime'>{0}</Value></Leq>" +
                                       "<Contains><FieldRef Name='Author' /><Value Type='UserMulti'>{2}</Value></Contains>" +
                                       "</And><Or><Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{0}</Value></Geq>" +
                                       "<Geq><FieldRef Name='EndDate' /><Value Type='DateTime'>{1}</Value></Geq>" +
                                       "</Or></And></Where></Query></View>", leaveEndDate, leaveStartDate, PName);

                                }
                                //TM  userAns = await GetAnsDetailsFromSharepoint(query, SpListName[0]);
                                //<Contains>< FieldRef Name = 'Author' />< Value Type = 'User' > krutika </ Value ></ Contains >
                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[0]);
                                if (userAns == "")
                                    userAns = "No data found for specific date";
                                break;
                            }


                    }
                    break;

                case "workingLocations":

                    switch (noOfLocation)
                    {
                        case 1:
                            {
                                //for single
                                query = string.Format("<View><Query><Where><And>" +
                                "<Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>{0}</Value></Eq>" +
                                "<And><Geq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{1}</Value></Geq>" +
                                "<Leq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{2}</Value></Leq>" +
                                "</And></And></Where></Query>" +
                                "<ViewFields><FieldRef Name='Date' /><FieldRef Name='Present' /><FieldRef Name='AssignedTo' /><FieldRef Name='WorkingLocation' /><FieldRef Name='Created' /></ViewFields><QueryOptions />" +
                                "</View>", workingLocation, leaveStartDate, leaveEndDate);
                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[1]);
                                break;
                            }
                        case 2:
                            {
                                if (workingLocation == "Client Location_Home")
                                {
                                    // for all 2 location (client/home)
                                    query = string.Format("<View><Query><Where><And>" +
                                        "<Leq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{0}</Value></Leq>" +
                                        "<And><Geq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{1}</Value></Geq>" +
                                        "<Or><Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Home</Value></Eq>" +
                                        "<Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Client Location</Value></Eq>" +
                                        "</Or></And></And></Where></Query>" +
                                        "<ViewFields><FieldRef Name='Date' /><FieldRef Name='Present' /><FieldRef Name='AssignedTo' /><FieldRef Name='WorkingLocation' /><FieldRef Name='Created' /></ViewFields><QueryOptions />" +
                                        "</View>", leaveEndDate, leaveStartDate);
                                }
                                if (workingLocation == "Client Location_Office")
                                {
                                    // for all 2 location (client/office)
                                    query = string.Format("<View><Query><Where><And>" +
                                    "<Leq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{0}</Value></Leq>" +
                                    "<And><Geq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{1}</Value></Geq>" +
                                    "<Or><Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Office</Value></Eq>" +
                                    "<Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Client Location</Value></Eq>" +
                                    "</Or></And></And></Where></Query>" +
                                    "<ViewFields><FieldRef Name='Date' /><FieldRef Name='Present' /><FieldRef Name='AssignedTo' /><FieldRef Name='WorkingLocation' /><FieldRef Name='Created' /></ViewFields><QueryOptions />" +
                                    "</View>", leaveEndDate, leaveStartDate);
                                }
                                if (workingLocation == "Office_Home")
                                {
                                    // for all 2 location (home/office)
                                    query = string.Format("<View><Query><Where><And>" +
                                    "<Leq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{0}</Value></Leq>" +
                                    "<And><Geq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{1}</Value></Geq>" +
                                    "<Or><Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Office</Value></Eq>" +
                                    "<Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Home</Value></Eq>" +
                                    "</Or></And></And></Where></Query>" +
                                    "<ViewFields><FieldRef Name='Date' /><FieldRef Name='Present' /><FieldRef Name='AssignedTo' /><FieldRef Name='WorkingLocation' /><FieldRef Name='Created' /></ViewFields><QueryOptions />" +
                                    "</View>", leaveEndDate, leaveStartDate);
                                }



                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[1]);
                                break;
                            }
                        case 3:
                            {
                                // for all 3 location
                                query = string.Format("<View><Query><Where><And>" +
                                    "<Leq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{0}</Value></Leq>" +
                                    "<And><Geq><FieldRef Name='Date' /><Value IncludeTimeValue='FALSE' Type='DateTime'>{1}</Value></Geq>" +
                                    "<Or><Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Office</Value></Eq>" +
                                    "<Or><Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Home</Value></Eq>" +
                                    "<Eq><FieldRef Name='WorkingLocation' /><Value Type='Choice'>Client Location</Value></Eq>" +
                                    "</Or></Or></And></And></Where></Query>" +
                                    "<ViewFields><FieldRef Name='Date' /><FieldRef Name='Present' /><FieldRef Name='AssignedTo' /><FieldRef Name='WorkingLocation' /><FieldRef Name='Created' /></ViewFields><QueryOptions />" +
                                    "</View>", leaveEndDate, leaveStartDate);

                                list1 = await GetAnsDetailsFromSharepoint(query, SpListName[1]);
                                break;
                            }
                    }
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
            staticListName = listName;
            using (ClientContext clientContext = new ClientContext("https://masycodasolutions.sharepoint.com/sites/OfficeMgmt"))
            {

                SecureString password = new SecureString();
                string pass = "TNVaq27606";
                foreach (char c in pass.ToCharArray()) password.AppendChar(c);
                clientContext.Credentials = new SharePointOnlineCredentials("Tejaswini@masycoda.com", pass);

                // TimeZoneInfo.ConvertTimeFromUtc(dt, TimeZoneInfo);
                Microsoft.SharePoint.Client.List UatList_ls = clientContext.Web.Lists.GetByTitle(listName);
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
                countOfLeave = items.Count;
                if (SpListName[0] == listName)
                {
                    if (items.Count != 0)
                    {
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
                            //String date1 = String.Concat(Convert.ToDateTime(items[i]["StartDate"]).AddHours(6).ToString("dd-MM-yyyy"),
                            // Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd-MM-yyyy"));
                            // Reason = items[i]["LeaveType"].ToString(),
                            String date1 = String.Concat(Convert.ToDateTime(items[i]["StartDate"]).AddHours(6).ToString("dd/MM/yyyy"), " ");
                            String date2 = String.Concat(date1, Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd/MM/yyyy"));
                            list1.Add(new Leave()
                            {
                                Author = childId_Value,
                                // StartDate = items[i]["StartDate"].ToString(),
                                //StartDate = Convert.ToDateTime(items[i]["StartDate"]).AddHours(6).ToString("dd-MM-yyyy"),
                                // EndDate = Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd-MM-yyyy"),
                                Reason = items[i]["LeaveType"].ToString(),
                                StartDate = date2,
                                Team = str1,
                                Manager = childId_Value1,
                            });


                        }

                    }

                    else
                    {
                        list1.Add(new Leave()
                        {
                            Author = "--",
                            // StartDate = items[i]["StartDate"].ToString(),
                            //StartDate = Convert.ToDateTime(items[i]["StartDate"]).AddHours(6).ToString("dd-MM-yyyy"),
                            // EndDate = Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd-MM-yyyy"),
                            Reason = "--",
                            StartDate = "--",
                            Team = "--",
                            Manager = "--",
                        });
                    }

                }
                //*******************************



                // attendance list
                else if (SpListName[1] == listName)
                {
                    if (items.Count != 0)
                    {



                        for (int i = 0; i < items.Count; i++)
                        {
                            //***Added first time
                            // text += "\n" + items[i]["StartDate"].ToString();
                            //****
                            String date1 = String.Concat(Convert.ToDateTime(items[i]["StartDate"]).AddHours(6).ToString("dd/MM/yyyy"), " ");
                            String date2 = String.Concat(date1, Convert.ToDateTime(items[i]["EndDate"]).AddHours(6).ToString("dd/MM/yyyy"));

                            var childIdField = items[i]["AssignedTo"] as FieldLookupValue;
                            var childId_Value = childIdField.LookupValue;



                            list1.Add(new Leave()
                            {
                                AssignedTo = childId_Value,

                                Date = Convert.ToDateTime(items[i]["Date"]).AddHours(6).ToString("dd-MM-yyyy"),
                                StartDate = date2,
                                WorkingLocation = items[i]["WorkingLocation"].ToString(),
                                Present = items[i]["Present"].ToString(),


                            });


                        }

                    }
                    else
                    {
                        list1.Add(new Leave()
                        {
                            AssignedTo = "--",

                            Date = "--",
                            WorkingLocation = "--",
                            Present = "--",
                            StartDate = "--",
                        });
                    }

                }


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
