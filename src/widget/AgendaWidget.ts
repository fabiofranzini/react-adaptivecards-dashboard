import { WidgetSize } from "@pnp/spfx-controls-react/lib/controls/dashboard/widget/IWidget";
import IAdaptiveCardWidget from "../models/IAdaptiveCardWidget";
import { IAdaptiveCardHostActionResult } from "@pnp/spfx-controls-react/lib/AdaptiveCardHost";

export const AgendaWidget: IAdaptiveCardWidget = {
    id: "agenda-widget",
    title: "Today's Agenda",
    size: WidgetSize.Double,
    adaptiveCard: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "TextBlock",
                "text": "Today's Agenda",
                "wrap": true,
                "style": "heading"
            },
            {
                "type": "ColumnSet",
                "horizontalAlignment": "center",
                "columns": [
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "ColumnSet",
                                "horizontalAlignment": "center",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://adaptivecards.io/content/LocationGreen_A.png",
                                                "altText": "Location A"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "**Redmond**",
                                                "wrap": true
                                            },
                                            {
                                                "type": "TextBlock",
                                                "spacing": "none",
                                                "text": "8a - 12:30p",
                                                "wrap": true
                                            }
                                        ],
                                        "width": "auto"
                                    }
                                ]
                            }
                        ],
                        "width": 1
                    },
                    {
                        "type": "Column",
                        "spacing": "large",
                        "separator": true,
                        "items": [
                            {
                                "type": "ColumnSet",
                                "horizontalAlignment": "center",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://messagecardplayground.azurewebsites.net/assets/LocationBlue_B.png",
                                                "altText": "Location B"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "**Bellevue**",
                                                "wrap": true
                                            },
                                            {
                                                "type": "TextBlock",
                                                "spacing": "none",
                                                "text": "12:30p - 3p",
                                                "wrap": true
                                            }
                                        ],
                                        "width": "auto"
                                    }
                                ]
                            }
                        ],
                        "width": 1
                    },
                    {
                        "type": "Column",
                        "spacing": "large",
                        "separator": true,
                        "items": [
                            {
                                "type": "ColumnSet",
                                "horizontalAlignment": "center",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://messagecardplayground.azurewebsites.net/assets/LocationRed_C.png",
                                                "altText": "Location C"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "**Seattle**",
                                                "wrap": true
                                            },
                                            {
                                                "type": "TextBlock",
                                                "spacing": "none",
                                                "text": "8p",
                                                "wrap": true
                                            }
                                        ],
                                        "width": "auto"
                                    }
                                ]
                            }
                        ],
                        "width": 1
                    }
                ]
            },
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "ColumnSet",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "horizontalAlignment": "left",
                                                "url": "https://messagecardplayground.azurewebsites.net/assets/Conflict.png",
                                                "altText": "Calendar conflict"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "spacing": "none",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "2:00 PM",
                                                "wrap": true
                                            }
                                        ],
                                        "width": "stretch"
                                    }
                                ]
                            },
                            {
                                "type": "TextBlock",
                                "spacing": "none",
                                "text": "1hr",
                                "isSubtle": true,
                                "wrap": true
                            }
                        ],
                        "width": "110px"
                    },
                    {
                        "type": "Column",
                        "backgroundImage": {
                            "url": "https://messagecardplayground.azurewebsites.net/assets/SmallVerticalLineGray.png",
                            "fillMode": "repeatVertically",
                            "horizontalAlignment": "center"
                        },
                        "items": [
                            {
                                "type": "Image",
                                "horizontalAlignment": "center",
                                "url": "https://adaptivecards.io/content/CircleGreen_coffee.png",
                                "altText": "Location A: Coffee"
                            }
                        ],
                        "width": "auto",
                        "spacing": "none"
                    },
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "**${Subject}**",
                                "wrap": true
                            },
                            {
                                "type": "ColumnSet",
                                "spacing": "none",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://messagecardplayground.azurewebsites.net/assets/location_gray.png",
                                                "altText": "Location"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "${Location.DisplayName}",
                                                "wrap": true
                                            }
                                        ],
                                        "width": "stretch"
                                    }
                                ]
                            },
                            {
                                "type": "ImageSet",
                                "spacing": "small",
                                "imageSize": "small",
                                "images": [
                                    {
                                        "type": "Image",
                                        "url": "https://messagecardplayground.azurewebsites.net/assets/person_w1.png",
                                        "size": "small",
                                        "altText": "Person with bangs"
                                    },
                                    {
                                        "type": "Image",
                                        "url": "https://messagecardplayground.azurewebsites.net/assets/person_m1.png",
                                        "size": "small",
                                        "altText": "Person with glasses and short hair"
                                    },
                                    {
                                        "type": "Image",
                                        "url": "https://messagecardplayground.azurewebsites.net/assets/person_w2.png",
                                        "size": "small",
                                        "altText": "Person smiling"
                                    }
                                ]
                            },
                            {
                                "type": "ColumnSet",
                                "spacing": "small",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://messagecardplayground.azurewebsites.net/assets/power_point.png",
                                                "altText": "Powerpoint presentation"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "**Contoso Brand Guidelines** shared by **Susan Metters**",
                                                "wrap": true
                                            }
                                        ],
                                        "width": "stretch"
                                    }
                                ]
                            }
                        ],
                        "width": 40
                    }
                ]
            },
            {
                "type": "ColumnSet",
                "spacing": "none",
                "columns": [
                    {
                        "type": "Column",
                        "width": "110px"
                    },
                    {
                        "type": "Column",
                        "backgroundImage": {
                            "url": "https://messagecardplayground.azurewebsites.net/assets/SmallVerticalLineGray.png",
                            "fillMode": "repeatVertically",
                            "horizontalAlignment": "center"
                        },
                        "items": [
                            {
                                "type": "Image",
                                "horizontalAlignment": "center",
                                "url": "https://messagecardplayground.azurewebsites.net/assets/Gray_Dot.png",
                                "altText": "Gray dot"
                            }
                        ],
                        "width": "auto",
                        "spacing": "none"
                    },
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "ColumnSet",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://messagecardplayground.azurewebsites.net/assets/car.png",
                                                "altText": "Travel by car"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "about 45 minutes",
                                                "isSubtle": true,
                                                "wrap": true
                                            }
                                        ],
                                        "width": "stretch"
                                    }
                                ]
                            }
                        ],
                        "width": 40
                    }
                ]
            },
            {
                "type": "ColumnSet",
                "spacing": "none",
                "columns": [
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "TextBlock",
                                "spacing": "none",
                                "text": "8:00 PM",
                                "wrap": true
                            },
                            {
                                "type": "TextBlock",
                                "spacing": "none",
                                "text": "1hr",
                                "isSubtle": true,
                                "wrap": true
                            }
                        ],
                        "width": "110px"
                    },
                    {
                        "type": "Column",
                        "backgroundImage": {
                            "url": "https://messagecardplayground.azurewebsites.net/assets/SmallVerticalLineGray.png",
                            "fillMode": "repeatVertically",
                            "horizontalAlignment": "center"
                        },
                        "items": [
                            {
                                "type": "Image",
                                "horizontalAlignment": "center",
                                "url": "https://messagecardplayground.azurewebsites.net/assets/CircleBlue_flight.png",
                                "altText": "Location B: Flight"
                            }
                        ],
                        "width": "auto",
                        "spacing": "none"
                    },
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "**Alaska Airlines AS1021 flight to Chicago**",
                                "wrap": true
                            },
                            {
                                "type": "ColumnSet",
                                "spacing": "none",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://messagecardplayground.azurewebsites.net/assets/location_gray.png",
                                                "altText": "Location"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Seattle Tacoma International Airport (17801 International Blvd, Seattle, WA, United States)",
                                                "wrap": true
                                            }
                                        ],
                                        "width": "stretch"
                                    }
                                ]
                            },
                            {
                                "type": "Image",
                                "url": "https://messagecardplayground.azurewebsites.net/assets/SeaTacMap.png",
                                "size": "stretch",
                                "altText": "Map of the Seattle-Tacoma airport"
                            }
                        ],
                        "width": 40
                    }
                ]
            }
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5"
    },
    data: {
        "@odata.context": "https://outlook.office.com/api/beta/$metadata#Me/Events/$entity",
        "@odata.id": "https://outlook.office.com/api/beta/Users('ddfcd489-628b-40d7-b48b-57002df800e5@1717622f-1d94-4d0c-9d74-709fad664b77')/Events('AAMkAGI2TG93AAA=')",
        "@odata.etag": "W/\"nfZyf7VcrEKLNoU37KWlkQAAA0x48w==\"",
        "Id": "AAMkAGI2TG93AAA=",
        "ChangeKey": "nfZyf7VcrEKLNoU37KWlkQAAA0x48w==",
        "Categories": [],
        "CreatedDateTime": "2014-10-19T23:13:47.3959685Z",
        "LastModifiedDateTime": "2014-10-19T23:13:47.6772234Z",
        "Subject": "Contoso Campaign Status Meeting",
        "BodyPreview": "Setting up some time to review the budget and planning on the Contoso Project",
        "Body": {
            "ContentType": "HTML",
            "Content": "\r\n\r\n\r\n\r\n\r\nSetting up some time to review the budget and planning on the Contoso Project\r\n\r\n\r\n"
        },
        "Importance": "Normal",
        "HasAttachments": false,
        "Start": {
            "DateTime": "2014-10-13T21:00:00",
            "TimeZone": "Pacific Standard Time"
        },
        "End": {
            "DateTime": "2014-10-13T22:00:00",
            "TimeZone": "Pacific Standard Time"
        },
        "Location": {
            "DisplayName": "Conf Room Bravern-2/9050",
            "Address": null
        },
        "ShowAs": "Busy",
        "IsAllDay": false,
        "IsCancelled": false,
        "IsOrganizer": true,
        "ResponseRequested": true,
        "Type": "SeriesMaster",
        "SeriesMasterId": null,
        "Attendees": [
            {
                "EmailAddress": {
                    "Address": "janets@a830edad9050849NDA1.onmicrosoft.com",
                    "Name": "Janet Schorr"
                },
                "Status": {
                    "Response": "None",
                    "Time": "0001-01-01T00:00:00Z"
                },
                "Type": "Required"
            },
            {
                "EmailAddress": {
                    "Address": "pavelb@a830edad9050849NDA1.onmicrosoft.com",
                    "Name": "Pavel Bansky"
                },
                "Status": {
                    "Response": "None",
                    "Time": "0001-01-01T00:00:00Z"
                },
                "Type": "Required"
            }
        ],
        "Recurrence": {
            "Pattern": {
                "Type": "Weekly",
                "Interval": 1,
                "Month": 0,
                "Index": "First",
                "FirstDayOfWeek": "Sunday",
                "DayOfMonth": 0,
                "DaysOfWeek": ["Monday"]
            },
            "RecurrenceTimeZone": "Pacific Standard Time",
            "Range": {
                "Type": "NoEnd",
                "StartDate": "2014-10-13",
                "EndDate": "2014-11-13",
                "NumberOfOccurrences": 0
            }
        },
        "OriginalEndTimeZone": "Pacific Standard Time",
        "OriginalStartTimeZone": "Pacific Standard Time",
        "Organizer": {
            "EmailAddress": {
                "Address": "alexd@a830edad9050849NDA1.onmicrosoft.com",
                "Name": "Alex D"
            },
            "OnlineMeetingUrl": null
        }
    },
    onAction: function (action: IAdaptiveCardHostActionResult): void {
        alert(`Action executed: ${JSON.stringify(action)}`);
        alert(`ContextData : ${this.contextData}`);
    }
}