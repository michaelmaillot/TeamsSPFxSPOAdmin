/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import {
  Button,
  Loader,
  ParticipantRemoveIcon,
  Table,
  TableRowProps,
} from "@fluentui/react-northstar";
import { getSP } from 'PnPJsConfig';

import { SPFI } from '@pnp/sp';
import { IExternalUser } from '@pnp/sp-admin/types';

interface IExternalUsersState {
  isLoading: boolean;
  externalUsers: any[];
}

export default class ExternalUsers extends React.Component<{}, IExternalUsersState> {  
  public constructor() {
    super({});

    this.state = {
      isLoading: true,
      externalUsers: [],
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadExternalUsers().catch((err) => { console.log(err) });

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
          <Table header={["Name", "Invited as", "Cross tenant?", "Created", "Actions"]} rows={this.state.externalUsers} />
        }
        </>
    );
  }

  private _loadExternalUsers = async (): Promise<void> => {
    // eslint-disable-next-line @typescript-eslint/typedef
    const sp: SPFI = getSP();
    const externalUsers = (await sp.admin.office365Tenant.getExternalUsers()).ExternalUserCollection
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
              content: 
                <Button
                  icon={<ParticipantRemoveIcon />}
                  iconOnly title="Remove user"
                  onClick={async (event: { stopPropagation: () => void; }, data: any) => {
                  event.stopPropagation();
                  await sp.admin.office365Tenant.removeExternalUsers([user.UniqueId]);
                }} />
            }
          ]
        } as TableRowProps
      });

    this.setState({
      externalUsers: externalUsers
    });
  }
}