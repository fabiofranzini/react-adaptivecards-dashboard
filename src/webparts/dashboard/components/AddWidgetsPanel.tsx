import * as React from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import IAdaptiveCardWidget from "../../../models/IAdaptiveCardWidget";
import { DefaultButton, Panel, PrimaryButton, SelectionMode, Stack } from '@fluentui/react';
import { ListView } from '@pnp/spfx-controls-react/lib/ListView';

export interface IAddWidgetsPanelProps {
    context: WebPartContext;
    avaibleWidgets: IAdaptiveCardWidget[];
    onDismiss: () => void;
    onSelect: (widgetIds: string[]) => void;
}

export default function AddWidgetsPanel(props: IAddWidgetsPanelProps): JSX.Element {

    const [selectedWidgetIds, setSelectedWidgetIds] = React.useState([]);

    return (
        <Panel
            headerText="Add Widgets"
            isOpen={true}
            onDismiss={props.onDismiss}
            closeButtonAriaLabel="Close"
            isFooterAtBottom={true}
            onRenderFooterContent={() => (
                <Stack horizontal tokens={{ childrenGap: 10 }}>
                    <PrimaryButton onClick={() => props?.onSelect(selectedWidgetIds)}>
                        Save
                    </PrimaryButton>
                    <DefaultButton onClick={() => props?.onDismiss()}>Cancel</DefaultButton>
                </Stack>
            )}>
            <ListView
                items={props.avaibleWidgets}
                viewFields={[
                    {
                        name: 'title',
                        displayName: 'Title',
                    }
                ]}
                iconFieldName="FileRef"
                compact={false}
                selectionMode={SelectionMode.multiple}
                selection={(items: IAdaptiveCardWidget[]) => {
                    setSelectedWidgetIds(items.map((item) => item.id));
                }}
                showFilter={false}
                stickyHeader={true} />
        </Panel>
    );
}