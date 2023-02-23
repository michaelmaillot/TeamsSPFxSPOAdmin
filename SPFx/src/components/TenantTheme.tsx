/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/typedef */
/* eslint-disable react/jsx-no-bind */
/* eslint-disable @microsoft/spfx/no-async-await */
import * as React from 'react';
import {
  Button,
  FormInput,
  AddIcon,
  Form,
  Header,
  Flex,
  FormButton,
  FormTextArea,
  Table,
  BanIcon,
  CheckmarkCircleIcon,
  TrashCanIcon,
  EditIcon,
  Loader,
} from "@fluentui/react-northstar";
import { getSP } from 'PnPJsConfig';

import { SPFI } from '@pnp/sp';
import { IThemeProperties } from '@pnp/sp-admin';
import { Panel, SwatchColorPicker } from '@fluentui/react';

export interface ITenantThemeProps {
}

interface ITenantThemeState {
  isLoading: boolean;
  isPanelOpen: boolean;
  tenantThemes: IThemeProperties[];
  newThemeName: string;
  newThemePalette: string;
  editThemeName: string;
  editThemePalette: string;
  themes: any[];
}

export default class TenantTheme extends React.Component<ITenantThemeProps, ITenantThemeState> {
  private _sp: SPFI;

  public constructor(props: ITenantThemeProps) {
    super(props);

    this.state = {
      isLoading: true,
      isPanelOpen: false,
      tenantThemes: [],
      newThemeName: "",
      newThemePalette: "",
      editThemeName: "",
      editThemePalette: "",
      themes: [],
    };

    this._sp = getSP();
  }

  public componentDidMount(): void {
    this._getTenantThemes().catch(err => console.log(err));
  }

  public render(): React.ReactElement<ITenantThemeProps> {
    const newThemePanelContent: JSX.Element =
      <>
        <Header>Add Theme</Header>
        <Form onSubmit={this._addTenantTheme}>
          <FormInput
            label="Theme name"
            name="themeName"
            id="theme-name"
            value={this.state.newThemeName}
            required
            autoComplete='off'
            onChange={(ev, data) => { this.setState({ newThemeName: data.value }) }}
            disabled={this.state.isLoading}
            styles={{ paddingBottom: "15px" }}
          />
          <FormTextArea
            label="Theme palette"
            name="themePalette"
            id="theme-palette"
            value={this.state.newThemePalette}
            required
            autoComplete='off'
            onChange={(ev, data) => { this.setState({ newThemePalette: data.value }) }}
            disabled={this.state.isLoading}
            styles={{ paddingBottom: "15px", "> textarea": { width: "192px", height: "400px" } }}
          />
          <Flex gap='gap.small' styles={{ position: "absolute", bottom: "0", paddingBottom: "15px" }}>
            <FormButton loading={this.state.isLoading} content="Apply" primary />
            <FormButton content="Cancel" type='button' secondary onClick={this._closePanel} disabled={this.state.isLoading} />
          </Flex>
        </Form>
      </>;

    const editThemePanelContent: JSX.Element =
      <>
        <Header>Edit Theme</Header>
        <Form onSubmit={this._updateTheme}>
          <FormInput
            label="Theme name"
            name="themeName"
            id="theme-name"
            value={this.state.editThemeName}
            disabled
            styles={{ paddingBottom: "15px" }}
          />
          <FormTextArea
            label="Theme palette"
            name="themePalette"
            id="theme-palette"
            value={this.state.editThemePalette}
            required
            autoComplete='off'
            onChange={(ev, data) => { this.setState({ editThemePalette: data.value }) }}
            disabled={this.state.isLoading}
            styles={{ paddingBottom: "15px", "> textarea": { width: "192px", height: "400px" } }}
          />
          <Flex gap='gap.small' styles={{ position: "absolute", bottom: "0", paddingBottom: "15px" }}>
            <FormButton loading={this.state.isLoading} content="Apply" primary />
            <FormButton content="Cancel" type='button' secondary onClick={this._closePanel} disabled={this.state.isLoading} />
          </Flex>
        </Form>
      </>;
      
    return (
      <>
        <Button icon={<AddIcon />} text content="Add new Theme" iconPosition='before' loading={this.state.isLoading} onClick={(ev) => { this.setState({ isPanelOpen: true, newThemePalette: "" }) }} />
        {this.state.isLoading
          ?
          <Loader label="Loading..." />
          :
          <Table header={["Name", "Palette", "Is inverted", "Actions"]} rows={this.state.themes} />
        }
        <Panel
          isOpen={this.state.isPanelOpen}
          closeButtonAriaLabel="Close"
          onDismiss={(ev) => this._closePanel()}>
          {this.state.editThemeName ? editThemePanelContent : newThemePanelContent}
        </Panel>
      </>
    );
  }

  private _addTenantTheme = async (): Promise<void> => {
    if (!this.state.isLoading) {
      try {
        this.setState({
          isLoading: true,
        });

        const pal = {
          "palette": JSON.parse(this.state.newThemePalette)
        };

        await this._sp.admin.office365Tenant.addTenantTheme(this.state.newThemeName, JSON.stringify(pal));

        await this._getTenantThemes();

        this.setState({
          newThemeName: "",
          newThemePalette: "",
          isPanelOpen: false,
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

  private _closePanel = (): void => {
    this.setState({
      isPanelOpen: false,
    });
  }

  private _getTenantThemes = async (): Promise<void> => {
    try {
      const themes: any[] = await this._sp.admin.office365Tenant.getAllTenantThemes()

      const tableThemes = themes.map((theme, index: number) => {

        const palette = theme.Palette.map((pal: any) => {
          return {
            id: pal.Key,
            label: pal.Key,
            color: pal.Value
          }
        });

        return {
          key: index + 1,
          items: [
            {
              content: theme.Name,
              styles: { maxWidth: "400px" }
            },
            {
              content: (
                <SwatchColorPicker
                  columnCount={8}
                  styles={{
                    root: {
                      tableLayout: "fixed",
                      width: "500px"
                    }
                  }}
                  cellShape={'circle'}
                  colorCells={palette}
                />),
              styles: { position: "relative", right: "8%" }
            },
            {
              content: (theme.IsInverted ? <CheckmarkCircleIcon /> : <BanIcon />),
              styles: { position: "relative", left: "3%" }
            },
            {
              content: (
                <Flex gap="gap.medium">
                  <Button iconOnly icon={<EditIcon />} title="Edit theme" onClick={(ev) => this._editTheme(theme)} />
                  <Button iconOnly icon={<TrashCanIcon />} title="Remove theme" onClick={(ev) => this._removeTheme(theme.Name)} />
                </Flex>),
            }
          ],
          styles: { minHeight: "150px" }
        }
      })

      this.setState({
        tenantThemes: themes,
        themes: tableThemes
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

  private _editTheme = (theme: any): void => {

    const editThemePalette = {} as any;
    for (const pal of theme.Palette) {
      editThemePalette[pal.Key] = pal.Value
    }

    console.log(JSON.stringify(editThemePalette));


    this.setState({
      isPanelOpen: true,
      editThemeName: theme.Name,
      editThemePalette: JSON.stringify(editThemePalette)
    })
  }

  private _updateTheme = async (): Promise<void> => {
    if (!this.state.isLoading) {
      try {
        this.setState({
          isLoading: true,
        });

        const pal = {
          "palette": JSON.parse(this.state.editThemePalette)
        };

        await this._sp.admin.office365Tenant.updateTenantTheme(this.state.editThemeName, JSON.stringify(pal));

        await this._getTenantThemes();

        this.setState({
          editThemeName: "",
          editThemePalette: "",
          isPanelOpen: false,
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

  private _removeTheme = async (themeName: string): Promise<void> => {
    try {
      this.setState({
        isLoading: true
      });
      await this._sp.admin.office365Tenant.deleteTenantTheme(themeName);
      await this._getTenantThemes();
    } catch (error) {
      console.log(error);
    }
    finally {
      this.setState({
        isLoading: false
      });
    }
  }
}