/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @microsoft/spfx/no-async-await */
/* eslint-disable react/jsx-no-bind */
import * as React from 'react';
import {
  Header,
  Flex,
  Form,
  FormInput,
  FormButton,
  FormCheckbox,
  Label,
  Divider
} from "@fluentui/react-northstar";

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Panel } from '@fluentui/react/lib/components/Panel/Panel';

export interface ICreateGroupProps {
  spHttpCtx: SPHttpClient;
  isPanelOpen: boolean;
  siteName: string;
  siteUrl: string;
  closePanel: () => void;
  closePanelSuccess: () => void;
}

interface ICreateGroupState {
  isLoading: boolean;
  newLockState: string;
  groupName: string;
  groupAlias: string;
  isGroupPublic: boolean;
}

export default class CreateGroup extends React.Component<ICreateGroupProps, ICreateGroupState> {
  public constructor(props: ICreateGroupProps) {
    super(props);

    this.state = {
      isLoading: false,
      newLockState: "",
      groupName: "",
      groupAlias: "",
      isGroupPublic: true,
    };
  }

  public render(): React.ReactElement<ICreateGroupProps> {
    return (
      <Panel
        styles={{ main: { width: "500px !important" } }} // TODO: get styles from Northstar theme to apply
        isOpen={this.props.isPanelOpen}
        closeButtonAriaLabel="Close"
        onDismiss={(_ev) => this.props.closePanel()}>
        <Header>Create Group</Header>
        <b>Site: </b><Label circular color="brand">{this.props.siteName}</Label>
        <Divider styles={{ paddingTop: "15px", paddingBottom: "15px" }} />
        <Form onSubmit={this._createSiteGroup} styles={{ justifyContent: "normal" }}>
          <FormInput
            label="Group name"
            name="groupName"
            id="group-name"
            value={this.state.groupName}
            required
            autoComplete='off'
            onChange={(_ev, data) => { this.setState({ groupName: data.value }) }}
            disabled={this.state.isLoading}
          />
          <FormInput
            label="Group alias"
            name="groupAlias"
            id="group-alias"
            value={this.state.groupAlias}
            required
            autoComplete='off'
            onChange={(_ev, data) => { this.setState({ groupAlias: data.value }) }}
            disabled={this.state.isLoading}
          />
          <FormCheckbox
            label="Is the group public?"
            checked={this.state.isGroupPublic}
            onChange={(_ev, data) => this.setState({ isGroupPublic: data.checked })}
            disabled={this.state.isLoading}
          />
          <Flex gap='gap.small' styles={{ position: "absolute", bottom: "0", paddingBottom: "15px" }}>
            <FormButton loading={this.state.isLoading} content="Apply" primary />
            <FormButton content="Cancel" type='button' secondary onClick={(_ev, _data) => this.props.closePanel()} disabled={this.state.isLoading} />
          </Flex>
        </Form>
      </Panel>
    );
  }

  private _createSiteGroup = async (_event: any, _data: any): Promise<void> => {
    if (!this.state.isLoading) {
      this.setState({
        isLoading: true,
      });

      try {
        const postURL: string = this.props.siteUrl + "/_api/GroupSiteManager/CreateGroupForSite";

        const httpClientOptions: ISPHttpClientOptions = {
          body: JSON.stringify({
            "alias": this.state.groupAlias,
            "displayName": this.state.groupName,
            "isPublic": this.state.isGroupPublic,
          }),
        };

        const newGroup: any = await this.props.spHttpCtx.post(
          postURL,
          SPHttpClient.configurations.v1,
          httpClientOptions)
          .then((response: SPHttpClientResponse): Promise<SPHttpClientResponse> => {
            return response.json();
          })
          .catch((err) => { throw err });

        console.log(newGroup);

        this.props.closePanelSuccess();
      } catch (error) {
        console.log(error);
      }
      finally {
        this.setState({
          isLoading: false,
        });
      }
    }
  }
}