import * as React from 'react';
import { IPnPDemoSiteDesignsProps } from './IPnPDemoSiteDesignsProps';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { sp, SiteScriptInfo, SiteDesignInfo } from '@pnp/sp';
import { autobind } from '@uifabric/utilities/lib';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { ISiteDesignRun, ISiteScriptActionStatus } from '@pnp/sp/src/sitedesigns';

export interface IPnPDemoSiteDesignsState {
  loading: boolean;
  scriptName: string;
  scriptJSON: string;
  scriptId: string;
  designName: string;
  designTemplate: string;
  designId: string;
  siteUrl: string;
  listTitle: string;
  scriptJSONList: string;
  // siteDesignHistory: string;
  siteDesignHistory: ISiteDesignRun[];
  allSiteScripts: SiteScriptInfo[];
  allSiteDesigns: SiteDesignInfo[];
}

export default class PnPDemoSiteDesigns extends React.Component<IPnPDemoSiteDesignsProps, IPnPDemoSiteDesignsState> {
  constructor(props) {
    super(props);
    this.state = this.initialState();
  }

  private initialState(): IPnPDemoSiteDesignsState {
    return {
      loading: false,
      scriptName: "",
      scriptJSON: "",
      scriptId: "",
      designName: "",
      designTemplate: "64",
      designId: "",
      siteUrl: "",
      allSiteScripts: [],
      allSiteDesigns: [],
      listTitle: "",
      scriptJSONList: "",
      siteDesignHistory: []
    };
  }

  public render(): React.ReactElement<IPnPDemoSiteDesignsProps> {
    let { loading, designTemplate } = this.state;
    console.log(designTemplate);
    const loadingSpinner: JSX.Element = <Spinner size={SpinnerSize.small} />;

    return (
      <div>
        <h3>1) Clear all site scripts and site designs</h3>
        <p>
          <PrimaryButton
            onClick={this.clearSiteDesigns}
            disabled={this.state.loading}
          >
            Clear all site scripts and site designs {loading ? loadingSpinner : null}
          </PrimaryButton>
        </p>

        <h3>2) Create a site script</h3>
        <p>
          <TextField label="Site Script name" onChanged={(val) => this.setState({ scriptName: val })} value={this.state.scriptName} />
          <TextField label="Site Script JSON" multiline rows={4} onChanged={(val) => this.setState({ scriptJSON: val })} value={this.state.scriptJSON} />
          <PrimaryButton onClick={this.createSiteScript} disabled={this.state.loading}>
            Create site script {loading ? loadingSpinner : null}
          </PrimaryButton>
          <p>Result Site Script ID: {this.state.scriptId}</p>
        </p>

        <h3>3) Create a site design (above created site script ID will be used)</h3>
        <p>
          <TextField label="Site Design name" onChanged={(val) => this.setState({ designName: val })} value={this.state.designName} />
          <ChoiceGroup
            options={[{ key: '64', text: 'Team Site', } as IChoiceGroupOption, { key: '68', text: 'Communication Site' } as IChoiceGroupOption]}
            onChange={(a?: any, option?: IChoiceGroupOption): void => this.setState({ ...this.state, designTemplate: option.key })}
            label="Site Template"
            selectedKey={designTemplate}
            required={true}
          />
          <PrimaryButton onClick={this.createSiteDesign} disabled={this.state.loading}>
            Create site design {loading ? loadingSpinner : null}
          </PrimaryButton>
          <p>Result Site Design ID: {this.state.designId}</p>
        </p>


        <h3>4) Apply site design on a site</h3>
        <p>
          <TextField label="Site URL" onChanged={(val) => this.setState({ siteUrl: val })} value={this.state.siteUrl} />
          <PrimaryButton onClick={this.applySiteDesign} disabled={this.state.loading}>
            Apply site design {loading ? loadingSpinner : null}
          </PrimaryButton>
          <PrimaryButton onClick={this.getHistoryRunOnSite} disabled={this.state.loading} style={{ paddingLeft: "10px" }}>
            Fetch Run History on site {loading ? loadingSpinner : null}
          </PrimaryButton>
          <ul>
            {this.state.siteDesignHistory.map(x => <li>{x.SiteDesignTitle} | RunID: {x.ID} | DesignID: {x.SiteDesignID}</li>)}
          </ul>
        </p>

        <h3>5) Get all site scripts</h3>
        <p>
          <PrimaryButton onClick={this.getAllSiteScripts} disabled={this.state.loading}>
            Get Site Scripts {loading ? loadingSpinner : null}
          </PrimaryButton>
          <ul>
            {this.state.allSiteScripts.map(x => <li>{x.Title} | {x.Id} | {x.Content}</li>)}
          </ul>
        </p>

        <h3>5) Get all site designs</h3>
        <p>
          <PrimaryButton onClick={this.getAllSiteDesigns} disabled={this.state.loading}>
            Get Site Designs {loading ? loadingSpinner : null}
          </PrimaryButton>
          <ul>
            {this.state.allSiteDesigns.map(x => <li>{x.Title} | {x.Id} | {x.WebTemplate}y
            </li>)}
          </ul>
        </p>

        <h3>6) Get site script from a list</h3>
        <p>
          <TextField label="List title" onChanged={(val) => this.setState({ listTitle: val })} value={this.state.listTitle} />
          <TextField label="List Script Result" multiline rows={10} value={this.state.scriptJSONList} />
          <PrimaryButton onClick={this.getSiteScriptFromList} disabled={this.state.loading}>
            Get Site Script From List {loading ? loadingSpinner : null}
          </PrimaryButton>
        </p>
      </div>
    );
  }

  @autobind
  private async getSiteScriptFromList() {
    let scriptJSON: string = "";
    this.setState({ loading: true });

    try {
      scriptJSON = await sp.web.lists.getByTitle(this.state.listTitle).getSiteScript();
    }
    catch (e) {
      alert(`Failed to fetch script from list ${this.state.listTitle}`);
    }

    this.setState({
      scriptJSONList: scriptJSON,
      loading: false
    });
  }

  @autobind
  private async clearSiteDesigns() {
    // Clear everything
    this.setState({ loading: true });

    const allSiteDesigns = await sp.siteDesigns.getSiteDesigns();
    for (let i = 0; i < allSiteDesigns.length; i++) {
      await sp.siteDesigns.deleteSiteDesign(allSiteDesigns[i].Id);
    }

    let allSiteScripts = await sp.siteScripts.getSiteScripts();
    for (let i = 0; i < allSiteScripts.length; i++) {
      await sp.siteScripts.deleteSiteScript(allSiteScripts[i].Id);
    }

    this.setState(this.initialState());
  }

  @autobind
  private async createSiteScript() {
    this.setState({ loading: true });

    const siteScript = await sp.siteScripts.createSiteScript(this.state.scriptName,
      "Description dummy..",
      JSON.parse(this.state.scriptJSON));

    const scId = siteScript.Id;
    this.setState({ loading: false, scriptId: scId });
  }

  @autobind
  private async createSiteDesign() {
    this.setState({ loading: true });
    // Create our site design with the above site script
    // WebTemplate: 64 Team site template, 68 Communication site template
    const createdSiteDesign = await sp.siteDesigns.createSiteDesign({
      Description: "PnPjs sample site design dummy...",
      SiteScriptIds: [this.state.scriptId],
      Title: this.state.designName,
      WebTemplate: this.state.designTemplate,
    });
    const sdId = createdSiteDesign.Id;
    this.setState({ loading: false, designId: sdId });
  }

  @autobind
  private async getHistoryRunOnSite() {
    this.setState({ loading: true });
    const result = await sp.siteDesigns.getSiteDesignRun(this.state.siteUrl);

    try {
      const runStatus: ISiteScriptActionStatus[] = await sp.siteDesigns
      .getSiteDesignRunStatus(this.state.siteUrl, result[0].ID);
      console.log(runStatus);
    }
    catch (ex) { }

    this.setState({
      loading: false,
      siteDesignHistory: result
    });
  }

  @autobind
  private async applySiteDesign() {
    this.setState({ loading: true });
    await sp.siteDesigns.applySiteDesign(this.state.designId, this.state.siteUrl);
    this.setState({ loading: false });
  }

  @autobind
  private async getAllSiteScripts() {
    this.setState({ loading: true });
    const allSiteScripts = await sp.siteScripts.getSiteScripts();
    // const siteScriptMetadata = await sp.siteScripts.getSiteScriptMetadata(scId);
    this.setState({ loading: false, allSiteScripts: allSiteScripts });
  }

  @autobind
  private async getAllSiteDesigns() {
    this.setState({ loading: true });
    const allSiteDesigns = await sp.siteDesigns.getSiteDesigns();
    this.setState({ loading: false, allSiteDesigns: allSiteDesigns });
  }
}
