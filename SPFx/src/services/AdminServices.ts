import { InjectHeaders, Caching } from "@pnp/queryable";
import { spfi, SPFI } from '@pnp/sp';
import { CopyFrom } from "@pnp/core";
import { getSPAdmin } from "PnPJsConfig";
import { ISearchBuilder, SearchQueryBuilder, SearchResults } from "@pnp/sp/search";
import Constants from "Constants";

export default class AdminServices {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public static async UpdateSiteProperties(siteUrl: string, updatedProperties: Record<string, any>): Promise<void> {
    const spAdmin: SPFI = getSPAdmin();

    const query: ISearchBuilder = SearchQueryBuilder().text("Path=" + siteUrl).selectProperties("SiteId");
    const results: SearchResults = await spAdmin.using(Caching()).search(query);

    await spfi(Constants.TENANT_ADMIN_URL).using(CopyFrom(spAdmin.site)).using(InjectHeaders({
      "Accept": "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-HTTP-Method": "MERGE"
    })).admin.tenant.call<void>("Sites('" + results.PrimarySearchResults[0].SiteId + "')", {
      "__metadata": {
        "type": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties"
      },
      ...updatedProperties
    });
  }
}