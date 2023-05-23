import * as React from 'react';
import { Dashboard as PnPDashboard } from '@pnp/spfx-controls-react/lib/Dashboard';
import { AdaptiveCardHost } from "@pnp/spfx-controls-react/lib/AdaptiveCardHost";
import { StaticCatalog } from '../../../widget/StaticCatalog';

import { WebPartContext } from "@microsoft/sp-webpart-base";
import IAdaptiveCardWidget from "../../../models/IAdaptiveCardWidget";
import AddWidgetsPanel from './AddWidgetsPanel';

import { AdaptiveCardDesignerHost } from "@pnp/spfx-controls-react/lib/AdaptiveCardDesignerHost";
import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

export interface IDashboardProps {
  context: WebPartContext;
  widgetIds: string[];

  canteenOrderList: string;
  emailColumn: string;
  entreeColumn: string;
  sideColumn: string;
  drinkColumn: string;

  title: string;
  onAddWidgets: (widgetIds: string[]) => void;
  onUpdateTitle: (title: string) => void;
  displayMode: DisplayMode;
}

export default function Dashboard(props: IDashboardProps): JSX.Element {

  const [cardsToRender, setCardsToRender] = React.useState([]);
  const [cardsToShow, setCardsToShow] = React.useState(StaticCatalog.filter((widget) => props.widgetIds?.some((id) => id === widget.id)));
  const [showAddWidget, setShowAddWidget] = React.useState(false);
  const [isInEditMode, setIsInEditMode] = React.useState(false);

  const isWellConfigured = React.useMemo(() => {
    return (props.canteenOrderList && props.emailColumn && props.entreeColumn && props.sideColumn && props.drinkColumn);
  }, [props.canteenOrderList, props.emailColumn, props.entreeColumn, props.sideColumn, props.drinkColumn]);

  React.useEffect(() => {
    const cardsWithContext = cardsToShow?.map((widget) => {
      switch (widget.id) {
        case 'canteen-order-widget': return {
          ...widget,
          contextData: {
            canteenOrderList: props.canteenOrderList,
            emailColumn: props.emailColumn,
            entreeColumn: props.entreeColumn,
            sideColumn: props.sideColumn,
            drinkColumn: props.drinkColumn,
            context: props.context
          }
        };
        default: return widget;
      }
    });
    setCardsToRender(cardsWithContext);
  }, [cardsToShow]);

  return (
    <>
      <div style={{ paddingLeft: 10 }}>
        <WebPartTitle displayMode={props.displayMode}
          title={props.title}
          updateProperty={props.onUpdateTitle} />
      </div>
      {!isWellConfigured && (
        <Placeholder iconName='Edit'
          iconText='Configure your web part'
          description='Please configure the web part!' />
      )}
      {isWellConfigured && (
        <>
          <PnPDashboard
            toolbarProps={props.displayMode === DisplayMode.Edit ? {
              actionGroups: {
                'group': {
                  'addAction': {
                    title: 'Add',
                    iconName: 'Add',
                    onClick: () => {
                      setIsInEditMode(false);
                      setShowAddWidget(true);
                    }
                  },
                  'edit': {
                    title: !isInEditMode ? 'Edit' : 'Stop Editing',
                    iconName: 'Edit',
                    onClick: () => {
                      setIsInEditMode(!isInEditMode);
                    }
                  },
                  'reset': {
                    title: 'Reset',
                    iconName: 'Reset',
                    onClick: () => {
                      setIsInEditMode(false);
                      setCardsToShow(StaticCatalog);
                    }
                  }
                }
              }
            } : undefined}
            widgets={
              cardsToRender?.map((card: IAdaptiveCardWidget) => {
                return {
                  id: card.title,
                  title: card.title,
                  size: card.size,
                  widgetActionGroup: props.displayMode === DisplayMode.Edit ? [
                    {
                      id: 'remove_widget',
                      title: 'Remove',
                      iconName: 'RemoveContent',
                      onClick: () => {
                        const cards = cardsToShow.filter((widget) => widget.id !== card.id);
                        setCardsToShow(cards);
                        props.onAddWidgets(cards.map((card) => card.id));
                      }
                    }
                  ] : undefined,
                  body: [
                    {
                      id: card.title,
                      title: card.title,
                      content: (
                        <>
                          {!isInEditMode &&
                            <div style={{ height: '100%', width: '100%', overflow: 'auto' }}>
                              <AdaptiveCardHost
                                card={card.adaptiveCard}
                                data={{ "$root": card.data }}
                                onInvokeAction={card.onAction.bind(card)}
                                onError={(error) => alert("Error: " + error.message)}
                                context={props.context as any}
                              />
                            </div>
                          }
                          {
                            isInEditMode && props.displayMode === DisplayMode.Edit &&
                            <AdaptiveCardDesignerHost
                              headerText="Adaptive Card Designer"
                              buttonText="Edit"
                              context={props.context as any}
                              onSave={(payload: object): void => { 
                                card.adaptiveCard = payload 
                              }}
                              addDefaultAdaptiveCardHostContainer={true}
                              card={card.adaptiveCard}
                              data={{ "$root": card.data }}
                            />
                          }
                        </>
                      )
                    }
                  ]
                }
              })
            } />
          {showAddWidget && (
            <AddWidgetsPanel
              context={props.context}
              avaibleWidgets={StaticCatalog.filter((widget) => !cardsToShow.some((card) => card.id === widget.id))}
              onDismiss={() => setShowAddWidget(false)}
              onSelect={(widgetIds: string[]) => {
                const addedWidgets = StaticCatalog.filter((widget) => widgetIds.some((id) => id === widget.id));
                const cards = [...cardsToShow, ...addedWidgets];
                setCardsToShow(cards);
                props.onAddWidgets(cards.map((card) => card.id));
                setShowAddWidget(false);
              }}
            />
          )}
        </>
      )}
    </>
  );
}
