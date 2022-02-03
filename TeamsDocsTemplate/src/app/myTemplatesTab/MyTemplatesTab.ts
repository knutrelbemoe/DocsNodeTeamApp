import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/myTemplatesTab/index.html")
@PreventIframe("/myTemplatesTab/config.html")
@PreventIframe("/myTemplatesTab/remove.html")
export class MyTemplatesTab {
}
