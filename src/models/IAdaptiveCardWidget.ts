import { IAdaptiveCardHostActionResult } from "@pnp/spfx-controls-react/lib/controls/adaptiveCardHost";
import { WidgetSize } from "@pnp/spfx-controls-react/lib/controls/dashboard/widget/IWidget";

export default interface IAdaptiveCardWidget {
    id: string;
    title: string;
    size: WidgetSize
    adaptiveCard: object;
    data: object;
    contextData?: object;
    onAction: (action: IAdaptiveCardHostActionResult) => void;
}