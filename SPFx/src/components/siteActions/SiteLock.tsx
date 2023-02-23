import * as React from 'react';
import {
  Button,
  LockIcon,
  ButtonProps,
  WandIcon,
  Tooltip,
} from "@fluentui/react-northstar";
import AdminServices from 'services/AdminServices';

export interface ISiteLockProps {
  lockState: string;
  siteUrl: string;
  disabled: boolean;
}

interface ISiteLockState {
  isLoading: boolean;
  newLockState: string;
}

export default class SiteLock extends React.Component<ISiteLockProps, ISiteLockState> {
  public constructor(props: ISiteLockProps) {
    super(props);

    this.state = {
      isLoading: false,
      newLockState: props.lockState
    };
  }

  public render(): React.ReactElement<ISiteLockProps> {
    return (
      <Tooltip
        trigger={
          <Button
            disabled={this.props.disabled}
            icon={this.state.newLockState === "Unlock" ? <LockIcon /> : <WandIcon />}
            loading={this.state.isLoading}
            iconOnly
            onClick={this._setSiteLockState} />}
        content={this.state.newLockState === "Unlock" ? "Lock the site" : "Unlock the site"} />
    );
  }

  // eslint-disable-next-line @microsoft/spfx/no-async-await
  private _setSiteLockState = async (event: React.SyntheticEvent<HTMLElement, Event>, _data: ButtonProps): Promise<void> => {

    event.stopPropagation();

    this.setState({
      isLoading: true,
    });
    try {
      const lockState: string = this.state.newLockState === "Unlock" ? "ReadOnly" : "Unlock";
      await AdminServices.UpdateSiteProperties(this.props.siteUrl, { "LockState": lockState });

      this.setState({
        newLockState: lockState
      });
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