import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { sp, ClientSidePage, ClientSideWebpart } from "@pnp/sp";
import { SPFetchClient } from "@pnp/nodejs";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    const webUrl = req.body && req.body.webUrl;
    if (!webUrl) {
        context.res = {
            status: 400,
            body: "Please pass a name on the query string or in the request body"
        };
    }
    else {
        sp.setup({
            sp: {
                fetchClientFactory: () => {
                    return new SPFetchClient(webUrl, "e5760b29-2b40-4cbd-bdf2-3f18cb5102db", "SECRET");
                },
            }
        });

        const w = await sp.web.select("Title", "Description", "ServerRelativeUrl").get();
        const page = await ClientSidePage.fromFile(sp.web.getFileByServerRelativeUrl(`${w.ServerRelativeUrl}/SitePages/Home.aspx`));
        const partDefs = await sp.web.getClientSideWebParts();
        // In my case I had to filter the id using uppercase & between brackets!
        const partDef = partDefs.filter(c => c.Id === "{420E9655-D63F-4CF5-8126-AD8A06BACCAE}");

        if (partDef.length < 1) {
            throw new Error("Could not find the web part");
        }

        const part = ClientSideWebpart.fromComponentDef(partDef[0]);
        part.setProperties<any>({
            description: `We set this property in our Azure Function! Current site title: ${w.Title}`,
        });

        page.sections = [];
        page.addSection().addControl(part);
        await page.save();
    }
};

export default httpTrigger;
