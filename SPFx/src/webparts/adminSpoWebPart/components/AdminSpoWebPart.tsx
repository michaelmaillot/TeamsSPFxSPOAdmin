/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @microsoft/spfx/no-async-await */
/* eslint-disable react/jsx-no-bind */
import * as React from 'react';
import { IAdminSpoProps } from './IAdminSpoProps';
import { SPFI } from "@pnp/sp";
import { IExternalUser } from "@pnp/sp-admin";
import "@pnp/sp-admin";
import "@pnp/sp/sites";
import "@pnp/sp/search";
import {
  teamsDarkTheme,
  teamsTheme,
  teamsHighContrastTheme,
  Provider,
  ThemePrepared,
  Button,
  RadioGroup,
  ShorthandCollection,
  RadioGroupItemProps,
  Divider,
  ParticipantRemoveIcon,
} from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { getSP } from 'PnPJsConfig';
import Cdn from 'components/Cdn';
import TenantTheme from 'components/TenantTheme';
import TenantSettings from 'components/TenantSettings';
import Sites from 'components/Sites';
import ExternalUsers from 'components/externalUsers/ExternalUsers';

interface IAdminSpoState {
  currentTheme: ThemePrepared<any>;
  externalUsers: any[];
  selectedMenu: Menu;
  isLoading: boolean;
}

enum Menu {
  Tenant,
  Sites,
  ExternalUsers,
  Themes,
  Cdn
}

export default class AdminSpoWebPart extends React.Component<IAdminSpoProps, IAdminSpoState> {
  private _sp: SPFI;
  private readonly _defaultMenuValue: Menu = Menu.Sites;
  private readonly _menu: ShorthandCollection<RadioGroupItemProps> = [
    {
      key: "sites",
      styles: { width: '100px' },
      checkedIndicator: <Button primary content="Sites" />,
      indicator: <Button secondary content="Sites" />,
      value: Menu.Sites
    },
    {
      key: "tenant",
      styles: { width: '100px' },
      checkedIndicator: <Button primary content="Tenant" />,
      indicator: <Button secondary content="Tenant" />,
      value: Menu.Tenant
    },
    {
      key: "cdn",
      styles: { width: '100px' },
      checkedIndicator: <Button primary content="CDN" />,
      indicator: <Button secondary content="CDN" />,
      value: Menu.Cdn
    },
    {
      key: "themes",
      styles: { width: '60px' },
      checkedIndicator: <Button primary content="Themes" />,
      indicator: <Button secondary content="Themes" />,
      value: Menu.Themes
    },
    {
      key: "externalUsers",
      styles: {
        "> div": {
          width: "auto"
        }
      },
      checkedIndicator: <Button primary content="External Users" />,
      indicator: <Button secondary content="External Users" />,
      value: Menu.ExternalUsers
    }
  ];

  public constructor(props: IAdminSpoProps) {
    super(props);

    this._sp = getSP();

    this.state = {
      currentTheme: teamsTheme,
      externalUsers: [],
      selectedMenu: this._defaultMenuValue,
      isLoading: true,
    };
  }

  public componentDidMount(): void {
    if (this.props.hasTeamsContext) {
      microsoftTeams.initialize();
      microsoftTeams.getContext((context) => {
        this._applyTheme(context.theme);
      });

      microsoftTeams.registerOnThemeChangeHandler((theme: string) => {
        this._applyTheme(theme);
      });
    }
    else {
      this._applyTheme();
    }
    this._loadAdmin()
      .then(() => {
        this.setState({
          isLoading: false,
        });
      })
      .catch(err => console.log(err));
  }

  private _applyTheme = (themeToApply?: string): void => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let newTheme: ThemePrepared<any>;

    switch (themeToApply) {
      case "default":
        newTheme = teamsTheme;
        break;

      case "dark":
        newTheme = teamsDarkTheme;
        break;

      case "contrast":
        newTheme = teamsHighContrastTheme;
        break;

      default:
        newTheme = teamsTheme;
        break;
    }

    this.setState({
      currentTheme: newTheme
    });
  }

  public render(): React.ReactElement<IAdminSpoProps> {
    return (
      <Provider theme={this.state.currentTheme}>
        <RadioGroup
          checkedValue={this.state.selectedMenu}
          defaultCheckedValue={this._defaultMenuValue}
          onCheckedValueChange={this._onCheckedValueChangeRadioMenu}
          styles={{ padding: '40px' }}
          items={this._menu} />
        <Divider styles={{ paddingBottom: "15px" }} />
        <Sites spHttpClient={this.props.context.spHttpClient} display={this.state.selectedMenu === Menu.Sites} />
        {this.state.selectedMenu === Menu.Tenant &&
          <TenantSettings graphClient={this.props.context.msGraphClientFactory} />
        }
        {this.state.selectedMenu === Menu.Cdn &&
          <Cdn />
        }
        {this.state.selectedMenu === Menu.Themes &&
          <TenantTheme />
        }
        {this.state.selectedMenu === Menu.ExternalUsers &&
          <ExternalUsers />
        }
      </Provider>
    );
  }

  private _onCheckedValueChangeRadioMenu = (_event: any, data: { value: Menu; }): void => {
    this.setState({ selectedMenu: data.value as Menu });
  }

  private _loadAdmin = async (): Promise<void> => {
    // eslint-disable-next-line @typescript-eslint/typedef
    const externalUsers = (await this._sp.admin.office365Tenant.getExternalUsers()).ExternalUserCollection
      .map((user: IExternalUser, index: number) => {
        return {
          key: index + 1,
          items: [
            {
              content: Math.random().toString(36).substring(2,7) + " " + Math.random().toString(36).substring(2,10)
            },
            {
              content: Math.random().toString(36).substring(2,7) + "@" + user.InvitedAs.split('@')[1]
            },
            {
              content: user.IsCrossTenant ? "yes" : "no"
            },
            {
              content: user.WhenCreated
            },
            {
              content: <Button icon={<ParticipantRemoveIcon />} iconOnly title="Remove user" onClick={async (event: { stopPropagation: () => void; }, data: any) => {
                event.stopPropagation();
                await this._sp.admin.office365Tenant.removeExternalUsers([user.UniqueId]);
              }} />
            }
          ]
        }
      });

    this.setState({
      externalUsers: externalUsers
    });
  }
}
