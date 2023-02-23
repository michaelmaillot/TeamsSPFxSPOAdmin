/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable react/jsx-no-bind */
/* eslint-disable @microsoft/spfx/no-async-await */
import * as React from 'react';
import {
  Button,
  Flex,
  Loader,
  Alert,
  Table,
  BanIcon,
  CheckmarkCircleIcon,
  ContactGroupIcon,
  Label,
  Segment,
  Input,
  InputProps,
  MoreIcon,
  Tooltip
} from "@fluentui/react-northstar";
import { getSP, getSPAdmin } from 'PnPJsConfig';
import { SPFI } from '@pnp/sp';
import {
  IPowerAppsEnvironment,
  ISPOWebTemplatesInfo,
  ITenantSitePropertiesInfo,
  PersonalSiteFilter
} from '@pnp/sp-admin';
import CreateGroup from './CreateGroup';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISearchBuilder, SearchQueryBuilder, SearchResults } from '@pnp/sp/search';
import { Caching } from "@pnp/queryable";
import { Link } from 'office-ui-fabric-react';
import SiteLock from './siteActions/SiteLock';
import SearchBar from './siteActions/SearchBar';
import BlockDownload from './siteActions/BlockDownload';
import MoreActions from './siteActions/MoreActions';
// import HomeSite from './siteActions/HomeSite';

interface ISitesProps {
  spHttpClient: SPHttpClient;
  display: boolean;
}

interface ISitesState {
  isLoading: boolean;
  sites: any[];
  initialSites: any[];
  sitePropsLoading: boolean;
  displayGrpPanel: boolean;
  selectedSite: Partial<ITenantSitePropertiesInfo>;
  displayAlertSuccess: boolean;
  homeSite: string;
  displayMorePanel: boolean;
}

class SPTemplates {
  static readonly POINTPUBLISHING_HUB = "pointpublishinghub#0";
  static readonly SEARCH_CENTER = "srchcen#0";
  static readonly PERSONNAL_SITE = "spsmsitehost#0";
  static readonly APP_CATALOG = "appcatalog#0";
  static readonly COMMUNICATION_SITE = "sitepagepublishing#0";
  static readonly PRIVATE_CHANNEL_SITE = "teamchannel#0";
  static readonly TEAM_SITE_CLASSIC = "sts#0";
  static readonly TEAM_SITE_NO_GROUP = "sts#3";
  static readonly TEAM_SITE = "group#0";
  static readonly REDIRECT_SITE = "redirectsite#0";
}

export default class Sites extends React.Component<ISitesProps, ISitesState> {
  private _spAdmin: SPFI;

  public constructor(props: ISitesProps) {
    super(props);

    this.state = {
      isLoading: true,
      sites: [],
      initialSites: [],
      sitePropsLoading: true,
      displayGrpPanel: false,
      selectedSite: null,
      displayAlertSuccess: false,
      homeSite: "",
      displayMorePanel: false,
    };

    this._spAdmin = getSPAdmin();
  }

  public async componentDidMount(): Promise<void> {
    await this._querySites().catch((err) => { console.log(err) });

    this.setState({
      isLoading: false,
    });
  }

  public render(): React.ReactElement<{}> {
    return (
      <div style={{ display: this.props.display ? "block" : "none" }}>
        <div style={{ minHeight: "30px" }}>
          {this.state.displayAlertSuccess &&
            <Alert
              success
              content={"Group created for the site " + this.state.selectedSite?.Title}
              dismissible
              dismissAction={{
                'aria-label': 'close',
                onClick: () => { this.setState({ displayAlertSuccess: false, selectedSite: null }) }
              }}
            />
          }
        </div>
        {this.state.isLoading
          ?
          <Loader label='Loading...' />
          :
          <Flex column>
            <Segment>
              <b>Home Site:</b> <a href={this.state.homeSite} target="_blank" rel="noreferrer"><Label circular color="brand">{this.state.homeSite}</Label></a>
            </Segment>
            <Segment>
              <Input placeholder="Search site..." autoComplete='off' clearable onChange={this._onChangeFilterSites} />
            </Segment>
            <Segment>
              <Table header={["Name", "Url", "Template", "Has group?", "App catalog enabled?", "Actions"]} rows={this.state.sites} />
            </Segment>
          </Flex>
        }
        <CreateGroup
          spHttpCtx={this.props.spHttpClient}
          isPanelOpen={this.state.displayGrpPanel}
          siteName={this.state.selectedSite?.Title}
          siteUrl={this.state.selectedSite?.Url}
          closePanel={() => this.setState({ displayGrpPanel: false, selectedSite: null })}
          closePanelSuccess={() => this.setState({ displayGrpPanel: false, displayAlertSuccess: true })} />
        <MoreActions
          spHttpCtx={this.props.spHttpClient}
          isPanelOpen={this.state.displayMorePanel}
          isHomeSite={this.state.homeSite === this.state.selectedSite?.Url || this.state.homeSite + "/" === this.state.selectedSite?.Url}
          isCommSite={this.state.selectedSite?.Template.toLowerCase() === SPTemplates.COMMUNICATION_SITE}
          site={this.state.selectedSite}
          closePanel={() => this.setState({ displayMorePanel: false, selectedSite: null })}
          closePanelSuccess={() => this.setState({ displayMorePanel: false, displayAlertSuccess: true })}
          refreshHomeSite={this._updateHomeSite} />
      </div>
    );
  }

  private _onChangeFilterSites = (_event: any, data: InputProps & {
    value: string;
  }): void => {
    if (data.value.length >= 3) {
      const filteredSites = this.state.initialSites.filter(site => {
        return site.items.find((itm: any) => {
          if (typeof itm.content === "string") {
            return itm.content.indexOf(data.value) > 0
          }
          else {
            return itm.content.props?.children?.indexOf(data.value) > 0
          }
        });
      }).filter(site => site !== undefined);

      this.setState({
        sites: filteredSites,
      });
    }
    else {
      this.setState({
        sites: this.state.initialSites,
      });
    }
  }

  private _querySites = async (): Promise<void> => {

    const tenantTemplates: ISPOWebTemplatesInfo = await this._spAdmin.admin.tenant.getSPOTenantAllWebTemplates();
    console.log(tenantTemplates);

    const homeSiteDetails: IPowerAppsEnvironment[] = await this._spAdmin.admin.tenant.getPowerAppsEnvironments();
    console.log(homeSiteDetails);

    const query: ISearchBuilder = SearchQueryBuilder().text("contentClass:STS_List_336").rowLimit(100);
    const results: SearchResults = await this._spAdmin.using(Caching()).search(query);

    console.log(results.PrimarySearchResults);

    await this._getHomeSite();

    const siteProps = await this._spAdmin.admin.tenant.getSitePropertiesFromSharePointByFilters({ IncludePersonalSite: PersonalSiteFilter.UseServerDefault, StartIndex: null, IncludeDetail: true });
    console.log(siteProps);
    const sitesRows = siteProps.map((site, index: number) => {

      const hasGroup: boolean = site.GroupId !== "00000000-0000-0000-0000-000000000000";
      const isGroupCreationEnabled: boolean = site.Template.toLowerCase().indexOf("sts") > -1 && !hasGroup;
      const appCatalog: boolean = results.PrimarySearchResults.some(catalog => catalog.ParentLink === site.Url);

      return {
        key: index + 1,
        items: [
          {
            content: site.Title ? site.Title : this._getSpecialSiteName(site.Template)
          },
          {
            content: (
              <Link href={site.Url} target="_blank" rel="noreferrer">{site.Url}</Link>
            ),
            truncateContent: true
          },
          {
            content: this._getTemplateDisplayName(site.Template)
          },
          {
            content: (
              hasGroup ? <CheckmarkCircleIcon /> : <BanIcon />
            ),
            styles: { left: "20px" }
          },
          {
            content: (
              appCatalog ? <CheckmarkCircleIcon /> : <BanIcon />
            ),
            styles: { left: "45px" }
          },
          {
            content: (
              <Flex gap='gap.medium'>
                <Tooltip
                  trigger={<Button
                    key={"groupCreation" + index}
                    icon={<ContactGroupIcon />}
                    iconOnly
                    disabled={!isGroupCreationEnabled} onClick={() => this.setState({ displayAlertSuccess: false, displayGrpPanel: true, selectedSite: site })} />}
                  content="Add a Microsoft 365 Group" />
                <SiteLock key={"siteLock" + index} disabled={this._isSiteSpecial(site.Template) || site.Url === this.state.homeSite} lockState={site.LockState} siteUrl={site.Url} />
                <SearchBar key={"searchBar" + index} disabled={this._isSiteSpecial(site.Template)} siteUrl={site.Url} />
                <BlockDownload key={"blockDownload" + index} disabled={this._isSiteSpecial(site.Template)} siteUrl={site.Url} blockDownloadEnabled={site.BlockDownloadPolicy} />
                <Tooltip
                  trigger={<Button
                    icon={<MoreIcon />}
                    iconOnly
                    disabled={this._isSiteSpecial(site.Template)}
                    onClick={() => this.setState({ displayAlertSuccess: false, displayMorePanel: true, selectedSite: site })} />}
                  content="Update more properties" />
              </Flex>
            ),
            styles: { position: "relative", right: "5%" }
          }
        ]
      }
    });

    console.log(sitesRows);
    this.setState({
      sites: sitesRows,
      initialSites: sitesRows,
    });
  }

  private _updateHomeSite = (siteUrl: string): void => {
    this.setState({
      homeSite: siteUrl,
    });
  }

  private _getHomeSite = async (): Promise<void> => {
    const getURL: string = (await getSP().site()).Url + "/_api/SPHSite/Details/";

    const currentHomeSite: any = await this.props.spHttpClient.get(
      getURL,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse): Promise<SPHttpClientResponse> => {
        return response.json();
      })
      .catch((err) => { throw err });

    console.log(currentHomeSite);

    this.setState({
      homeSite: currentHomeSite.Url, //+ "/"
    });
  }

  private _isSiteSpecial = (templateId: string): boolean => {
    return [
      SPTemplates.SEARCH_CENTER,
      SPTemplates.PERSONNAL_SITE,
      SPTemplates.POINTPUBLISHING_HUB,
      SPTemplates.APP_CATALOG,
      SPTemplates.REDIRECT_SITE
    ].some(val => val === templateId.toLowerCase());
  }

  private _getSpecialSiteName = (templateId: string): string => {
    let siteName: string;

    switch (templateId.toLowerCase()) {
      case SPTemplates.SEARCH_CENTER: {
        siteName = "Search Center";
        break;
      }

      case SPTemplates.PERSONNAL_SITE: {
        siteName = "Personnal Site";
        break;
      }

      default:
        siteName = "";
        break;
    }

    return siteName;
  }

  private _getTemplateDisplayName = (templateId: string): string => {
    let templateName: string;
    switch (templateId.toLowerCase()) {
      case SPTemplates.TEAM_SITE_NO_GROUP:
      case SPTemplates.TEAM_SITE:
      case SPTemplates.TEAM_SITE_CLASSIC: {
        templateName = "Team Site" + (templateId.toLowerCase() === SPTemplates.TEAM_SITE_CLASSIC ? " (classic experience)" : "");
        break;
      }

      case SPTemplates.POINTPUBLISHING_HUB: {
        templateName = "PointPublishing Hub";
        break;
      }

      case SPTemplates.COMMUNICATION_SITE: {
        templateName = "Communication Site";
        break;
      }

      case SPTemplates.PRIVATE_CHANNEL_SITE: {
        templateName = "Private Channel Site";
        break;
      }

      case SPTemplates.SEARCH_CENTER: {
        templateName = "Enterprise Search Center";
        break;
      }

      case SPTemplates.PERSONNAL_SITE: {
        templateName = "Personnal Site";
        break;
      }

      case SPTemplates.APP_CATALOG: {
        templateName = "Tenant App Catalog";
        break;
      }

      default: {
        templateName = templateId;
        break;
      }
    }

    return templateName;
  }
}