import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/swoopAnalyticsTab/index.html")
@PreventIframe("/swoopAnalyticsTab/config.html")
@PreventIframe("/swoopAnalyticsTab/remove.html")
export class SwoopAnalyticsTab {
}
