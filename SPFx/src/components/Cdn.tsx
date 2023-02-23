/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable react/jsx-no-bind */
/* eslint-disable @microsoft/spfx/no-async-await */
import * as React from 'react';
import {
  Button,
  Checkbox,
  Flex,
  Form,
  FormInput,
  FormButton,
  Loader,
  CheckboxProps,
  List,
  Divider,
  AddIcon,
  Text,
  CloseIcon,
  ComponentEventHandler,
  ButtonProps
} from "@fluentui/react-northstar";
import { getSP } from 'PnPJsConfig';
import { SPFI } from '@pnp/sp';
import { SPOTenantCdnType } from '@pnp/sp-admin';

interface ICdnState {
  isLoading: boolean;
  cdnPublicEnabled: boolean;
  cdnPublicOrigins: any[];
  cdnPublicNewOrigin: string;
  cdnPrivateEnabled: boolean;
  cdnPrivateOrigins: any[];
  cdnPrivateNewOrigin: string;
}

export default class Cdn extends React.Component<{}, ICdnState> {
  private _sp: SPFI;

  public constructor() {
    super({});

    this.state = {
      isLoading: true,
      cdnPublicEnabled: false,
      cdnPublicOrigins: [],
      cdnPublicNewOrigin: "",
      cdnPrivateEnabled: false,
      cdnPrivateOrigins: [],
      cdnPrivateNewOrigin: "",
    };

    this._sp = getSP();
  }

  public async componentDidMount(): Promise<void> {
    await this._getCdnData().catch((err) => { console.log(err) });

    this.setState({
      isLoading: false,
    });
  }

  public render(): React.ReactElement<{}> {
    return (
      <>
        {this.state.isLoading
          ?
          <Loader label='Loading...' />
          :
          <Flex styles={{ paddingLeft: "15px" }} gap='gap.large'>
            <Flex column gap="gap.small">
              <Checkbox
                label="Public CDN enabled?"
                defaultChecked={this.state.cdnPublicEnabled}
                checked={this.state.cdnPublicEnabled}
                toggle
                onClick={this._setCdnPublicEnabled} />
              <Divider />
              <Text styles={{ textAlign: "center", fontWeight: "800", minHeight: '20px' }} content="Public origins" />
              <Form onSubmit={this._addCdnPublicOrigin}>
                <Flex gap='gap.small'>
                  <FormInput
                    required
                    placeholder="New public origin..."
                    autoComplete='off'
                    value={this.state.cdnPublicNewOrigin}
                    disabled={!this.state.cdnPublicEnabled}
                    onChange={(ev, data) => { this.setState({ cdnPublicNewOrigin: data.value }) }} />
                  <FormButton disabled={!this.state.cdnPublicEnabled} icon={<AddIcon />} iconOnly title="Add new public origin" />
                </Flex>
              </Form>
              <List styles={{ minHeight: "90px" }} items={this.state.cdnPublicOrigins} />
            </Flex>
            <Flex column gap="gap.small">
              <Checkbox
                label="Private CDN enabled?"
                defaultChecked={this.state.cdnPrivateEnabled}
                checked={this.state.cdnPrivateEnabled}
                toggle
                onClick={this._setCdnPrivateEnabled} />
              <Divider />
              <Text styles={{ textAlign: "center", fontWeight: "800", minHeight: '20px' }} content="Private origins" />
              <Form onSubmit={this._addCdnPrivateOrigin}>
                <Flex gap='gap.small'>
                  <FormInput
                    required
                    placeholder="New private origin..."
                    autoComplete='off'
                    value={this.state.cdnPrivateNewOrigin}
                    disabled={!this.state.cdnPrivateEnabled}
                    onChange={(ev, data) => { this.setState({ cdnPrivateNewOrigin: data.value }) }} />
                  <FormButton disabled={!this.state.cdnPrivateEnabled} icon={<AddIcon />} iconOnly title="Add new private origin" />
                </Flex>
              </Form>
              <List styles={{ minHeight: "90px" }} items={this.state.cdnPrivateOrigins} />
            </Flex>
          </Flex>
        }
      </>
    );
  }

  private _addCdnPublicOrigin = async (_event: React.SyntheticEvent<HTMLElement, Event>): Promise<void> => {
    try {
      await this._sp.admin.office365Tenant.call<void>(`AddTenantCdnOrigin`, {
        "cdnType": SPOTenantCdnType.Public,
        "originUrl": this.state.cdnPublicNewOrigin
      });

      const publicCdnOrigins: any[] = await this._getCdnOriginsList(SPOTenantCdnType.Public);

      this.setState({
        cdnPublicOrigins: publicCdnOrigins,
        cdnPublicNewOrigin: "",
      })
    } catch (error) {
      console.log(error);
    }
  }

  private _addCdnPrivateOrigin = async (_event: React.SyntheticEvent<HTMLElement, Event>): Promise<void> => {
    try {
      await this._sp.admin.office365Tenant.call<void>(`AddTenantCdnOrigin()`, {
        "cdnType": SPOTenantCdnType.Private,
        "originUrl": this.state.cdnPrivateNewOrigin
      });

      const privateCdnOrigins: any[] = await this._getCdnOriginsList(SPOTenantCdnType.Private);

      this.setState({
        cdnPublicOrigins: privateCdnOrigins,
        cdnPrivateNewOrigin: "",
      })
    } catch (error) {
      console.log(error);
    }
  }

  private _setCdnPublicEnabled = async (_event: React.SyntheticEvent<HTMLElement, Event>, data: CheckboxProps): Promise<void> => {
    let publicCdnOrigins: any[] = [...this.state.cdnPublicOrigins];

    try {
      this.setState({
        isLoading: true,
      });

      await this._sp.admin.office365Tenant.call<void>(`SetTenantCdnEnabled()`, {
        "cdnType": SPOTenantCdnType.Public,
        "isEnabled": data.checked
      });

      publicCdnOrigins = await this._getCdnOriginsList(SPOTenantCdnType.Public);
    } catch (error) {
      console.log(error);
    }
    finally {
      this.setState({
        isLoading: false,
        cdnPublicEnabled: data.checked as boolean,
        cdnPublicOrigins: publicCdnOrigins,
      });
    }
  }

  private _setCdnPrivateEnabled = async (_event: React.SyntheticEvent<HTMLElement, Event>, data: CheckboxProps): Promise<void> => {
    let privateCdnOrigins: any[] = [...this.state.cdnPrivateOrigins];
    
    try {

      this.setState({
        isLoading: true,
      });

      await this._sp.admin.office365Tenant.call<void>(`SetTenantCdnEnabled()`, {
        "cdnType": SPOTenantCdnType.Private,
        "isEnabled": data.checked
      });

      privateCdnOrigins = await this._getCdnOriginsList(SPOTenantCdnType.Private);
    } catch (error) {
      console.log(error);
    }
    finally {
      this.setState({
        isLoading: false,
        cdnPrivateEnabled: data.checked as boolean,
        cdnPrivateOrigins: privateCdnOrigins
      });
    }
  }

  private _getCdnData = async (): Promise<void> => {
    const publicCdnEnabled: boolean = await this._sp.admin.office365Tenant.call<boolean>(`GetTenantCdnEnabled(${SPOTenantCdnType.Public})`, {});
    const privateCdnEnabled: boolean = await this._sp.admin.office365Tenant.call<boolean>(`GetTenantCdnEnabled(${SPOTenantCdnType.Private})`, {});

    const publicCdnOrigins: any[] = await this._getCdnOriginsList(SPOTenantCdnType.Public);
    const privateCdnOrigins: any[] = await this._getCdnOriginsList(SPOTenantCdnType.Private);

    this.setState({
      cdnPublicEnabled: publicCdnEnabled,
      cdnPublicOrigins: publicCdnOrigins,
      cdnPrivateEnabled: privateCdnEnabled,
      cdnPrivateOrigins: privateCdnOrigins
    });
  }

  private _getCdnOriginsList = async (cdnType: SPOTenantCdnType): Promise<any[]> => {
    const cdnOrigins: string[] = await this._sp.admin.office365Tenant.call<string[]>(`GetTenantCdnOrigins(${cdnType})`, {});
    const cdnRemovalMethod: ComponentEventHandler<ButtonProps> = (cdnType === SPOTenantCdnType.Public ? this._removePublicCdnOrigin : this._removePrivateCdnOrigin);
    const cdnToggleState: boolean = (cdnType === SPOTenantCdnType.Public ? this.state.cdnPublicEnabled : this.state.cdnPrivateEnabled);

    return cdnOrigins.map(val => {
      return {
        content: (
          <Flex gap="gap.small">
            <Text styles={{ marginTop: "8px" }} content={val} />
            <Button icon={<CloseIcon />} disabled={cdnToggleState} text iconOnly title="Remove origin" data-cdn-origin={val.replace('(configuration pending)', '').trim()} onClick={cdnRemovalMethod} />
          </Flex>
        )
      }
    });
  }

  private _removePublicCdnOrigin = async (event: React.SyntheticEvent<HTMLElement, Event>): Promise<void> => {
    try {
      const cdnOriginToRemove: string = event.currentTarget.getAttribute("data-cdn-origin");

      await this._sp.admin.office365Tenant.call<void>(`RemoveTenantCdnOrigin()`, {
        "cdnType": SPOTenantCdnType.Public,
        "originUrl": cdnOriginToRemove
      });

      const cdnOrigins: string[] = await this._getCdnOriginsList(SPOTenantCdnType.Public);

      this.setState({
        cdnPublicOrigins: cdnOrigins
      })
    } catch (error) {
      console.log(error);
    }
  }

  private _removePrivateCdnOrigin = async (event: React.SyntheticEvent<HTMLElement, Event>): Promise<void> => {
    try {
      const cdnOriginToRemove: string = event.currentTarget.getAttribute("data-cdn-origin");

      await this._sp.admin.office365Tenant.call<void>(`RemoveTenantCdnOrigin()`, {
        "cdnType": SPOTenantCdnType.Private,
        "originUrl": cdnOriginToRemove
      });

      const cdnOrigins: string[] = await this._getCdnOriginsList(SPOTenantCdnType.Private);

      this.setState({
        cdnPrivateOrigins: cdnOrigins
      })
    } catch (error) {
      console.log(error);
    }
  }
}