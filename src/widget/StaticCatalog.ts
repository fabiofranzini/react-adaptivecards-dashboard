import IAdaptiveCardWidget from "../models/IAdaptiveCardWidget";
import { AgendaWidget } from "./AgendaWidget";
import { CanteenOrderWidget } from "./CanteenOrderWidget";
import { WeatherWidget } from "./WeatherWidget";

export const StaticCatalog: Array<IAdaptiveCardWidget> = [
    AgendaWidget,
    CanteenOrderWidget,
    WeatherWidget
];