import { WidgetSize } from "@pnp/spfx-controls-react/lib/controls/dashboard/widget/IWidget";
import IAdaptiveCardWidget from "../models/IAdaptiveCardWidget";
import { IAdaptiveCardHostActionResult } from "@pnp/spfx-controls-react/lib/controls/adaptiveCardHost";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { spfi, SPFx as spSPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";

export interface ICanteenOrderWidgetContext {
    canteenOrderList: string;
    emailColumn: string;
    entreeColumn: string;
    sideColumn: string;
    drinkColumn: string;
    context: WebPartContext;
}

export const CanteenOrderWidget: IAdaptiveCardWidget = {
    id: "canteen-order-widget",
    title: "Canteen Order",
    size: WidgetSize.Single,
    adaptiveCard: {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {
                "type": "TextBlock",
                "text": "${FormInfo.title}",
                "size": "large",
                "wrap": true,
                "weight": "bolder",
                "style": "heading"
            },
            {
                "type": "Input.ChoiceSet",
                "id": "EntreeSelectVal",
                "label": "${Order.questions[0].question}",
                "style": "filtered",
                "isRequired": true,
                "errorMessage": "This is a required input",
                "placeholder": "Please choose",
                "choices": [
                    {
                        "$data": "${Order.questions[0].items}",
                        "title": "${choice}",
                        "value": "${value}"
                    }
                ]
            },
            {
                "type": "Input.ChoiceSet",
                "id": "SideVal",
                "label": "${Order.questions[1].question}",
                "style": "filtered",
                "isRequired": true,
                "errorMessage": "This is a required input",
                "placeholder": "Please choose",
                "choices": [
                    {
                        "$data": "${Order.questions[1].items}",
                        "title": "${choice}",
                        "value": "${value}"
                    }
                ]
            },
            {
                "type": "Input.ChoiceSet",
                "id": "DrinkVal",
                "label": "${Order.questions[2].question}",
                "style": "filtered",
                "isRequired": true,
                "errorMessage": "This is a required input",
                "placeholder": "Please choose",
                "choices": [
                    {
                        "$data": "${Order.questions[2].items}",
                        "title": "${choice}",
                        "value": "${value}"
                    }
                ]
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Place Order"
            }
        ]
    }
    ,
    data: {
        "FormInfo": {
            "title": "Canteen Order Form"
        },
        "Order": {
            "questions": [
                {
                    "question": "Which entree would you like?",
                    "items": [
                        {
                            "choice": "Steak",
                            "value": "Steak"
                        },
                        {
                            "choice": "Chicken",
                            "value": "Chicken"
                        },
                        {
                            "choice": "Tofu",
                            "value": "Tofu"
                        }
                    ]
                },
                {
                    "question": "Which side would you like?",
                    "items": [
                        {
                            "choice": "Baked Potato",
                            "value": "Baked Potato"
                        },
                        {
                            "choice": "Rice",
                            "value": "Rice"
                        },
                        {
                            "choice": "Vegetables",
                            "value": "Vegetables"
                        },
                        {
                            "choice": "Noodles",
                            "value": "Noodles"
                        },
                        {
                            "choice": "No Side",
                            "value": "No Side"
                        }
                    ]
                },
                {
                    "question": "Which drink would you like?",
                    "items": [
                        {
                            "choice": "Water",
                            "value": "Water"
                        },
                        {
                            "choice": "Soft Drink",
                            "value": "Soft Drink"
                        },
                        {
                            "choice": "Coffee",
                            "value": "Coffee"
                        },
                        {
                            "choice": "Natural Juice",
                            "value": "Natural Juice"
                        },
                        {
                            "choice": "No Drink",
                            "value": "No Drink"
                        }
                    ]
                }
            ]
        }
    },
    onAction: function (action: IAdaptiveCardHostActionResult): void {
        const widgetContext: ICanteenOrderWidgetContext = this.contextData as ICanteenOrderWidgetContext;
        const sp = spfi().using(spSPFx(widgetContext.context));

        (async () => {
            try {
                const list = sp.web.lists.getById(widgetContext.canteenOrderList);
                let itemToAdd: { [key: string]: string } = {}
                itemToAdd[widgetContext.emailColumn] = widgetContext.context.pageContext.user.email;
                itemToAdd[widgetContext.entreeColumn] = (action.data as any)?.EntreeSelectVal;
                itemToAdd[widgetContext.sideColumn] = (action.data as any)?.SideVal;
                itemToAdd[widgetContext.drinkColumn] = (action.data as any)?.DrinkVal;
                const item: IItemAddResult = await list.items.add(itemToAdd);
                console.log(`Added item: ${JSON.stringify(item)}`);
                alert(`Order Added`);
            } catch (error) {
                console.error(error);
                alert(`An error has occurred, please try again`);
            }
        })();
    }
}

