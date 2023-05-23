import { WidgetSize } from "@pnp/spfx-controls-react/lib/controls/dashboard/widget/IWidget";
import IAdaptiveCardWidget from "../models/IAdaptiveCardWidget";
import { IAdaptiveCardHostActionResult } from "@pnp/spfx-controls-react/lib/controls/adaptiveCardHost";

export const WeatherWidget: IAdaptiveCardWidget = {
    id: "weather-widget",
    title: "Weather",
    size: WidgetSize.Single,
    adaptiveCard: {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "speak": "The forecast for Seattle ${formatEpoch(dt, 'MMMM d')} is mostly clear with a High of ${formatNumber((main.temp_max - 273) * 9 / 5 + 32, 0)} degrees and Low of ${formatNumber((main.temp_min - 273) * 9 / 5 + 32, 0)} degrees",
        "body": [
            {
                "type": "TextBlock",
                "text": "${name}, WA",
                "size": "large",
                "isSubtle": true,
                "wrap": true,
                "style": "heading"
            },
            {
                "type": "TextBlock",
                "text": "{{DATE(${formatEpoch(dt, 'yyyy-MM-ddTHH:mm:ssZ')}, SHORT)}} {{TIME(${formatEpoch(dt, 'yyyy-MM-ddTHH:mm:ssZ')})}}",
                "spacing": "none",
                "wrap": true
            },
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "width": "auto",
                        "items": [
                            {
                                "type": "Image",
                                "url": "https://messagecardplayground.azurewebsites.net/assets/Mostly%20Cloudy-Square.png",
                                "size": "small",
                                "altText": "Mostly cloudy weather"
                            }
                        ]
                    },
                    {
                        "type": "Column",
                        "width": "auto",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "${formatNumber((main.temp - 273) * 9 / 5 + 32, 0)}",
                                "size": "extraLarge",
                                "spacing": "none",
                                "wrap": true
                            }
                        ]
                    },
                    {
                        "type": "Column",
                        "width": "stretch",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "°F",
                                "weight": "bolder",
                                "spacing": "small",
                                "wrap": true
                            }
                        ]
                    },
                    {
                        "type": "Column",
                        "width": "stretch",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "Hi ${formatNumber((main.temp_max - 273) * 9 / 5 + 32, 0)}",
                                "horizontalAlignment": "left",
                                "wrap": true
                            },
                            {
                                "type": "TextBlock",
                                "text": "Lo ${formatNumber((main.temp_min - 273) * 9 / 5 + 32, 0)}",
                                "horizontalAlignment": "left",
                                "spacing": "none",
                                "wrap": true
                            }
                        ]
                    }
                ]
            }
        ]
    },
    data: {
        "coord": {
            "lon": -122.12,
            "lat": 47.67
        },
        "weather": [
            {
                "id": 802,
                "main": "Clouds",
                "description": "scattered clouds",
                "icon": "03n"
            }
        ],
        "base": "stations",
        "main": {
            "temp": 281.05,
            "pressure": 1022,
            "humidity": 81,
            "temp_min": 278.15,
            "temp_max": 283.15
        },
        "visibility": 16093,
        "wind": {
            "speed": 4.1,
            "deg": 360
        },
        "rain": {},
        "clouds": {
            "all": 40
        },
        "dt": 1572920478,
        "sys": {
            "type": 1,
            "id": 5798,
            "country": "US",
            "sunrise": 1572879421,
            "sunset": 1572914822
        },
        "timezone": -28800,
        "id": 5808079,
        "name": "Redmond",
        "cod": 200
    },
    onAction: function (action: IAdaptiveCardHostActionResult): void {
        alert(`Action executed: ${JSON.stringify(action)}`);
    }
}