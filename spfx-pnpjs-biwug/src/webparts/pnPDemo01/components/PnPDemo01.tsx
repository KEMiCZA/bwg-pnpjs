import * as React from 'react';
import { IPnPDemo01Props } from './IPnPDemo01Props';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

import { sp } from '@pnp/sp';

export interface Demo01State {
  loading: boolean;
  webInformation: any;
  subsites: any[];
  fields: any[];
  loadingCreatingItems: boolean;
  itemsCreationPerformance?: number;
  itemsCreationBatchPerformance?: number;
}

export default class PnPDemo01 extends React.Component<IPnPDemo01Props, Demo01State> {
  private readonly _itemscount = 50;
  private readonly _pnpDemoList01Name = "PnPDemo01";

  constructor(props) {
    super(props);

    this.state = {
      loading: true,
      webInformation: {},
      subsites: [],
      fields: [],
      loadingCreatingItems: false
    };

    this.loadWebInformation = this.loadWebInformation.bind(this);
    this._createItemsUsingBatching = this._createItemsUsingBatching.bind(this);
    this._createItems = this._createItems.bind(this);
  }

  public componentDidMount() {
    this.loadWebInformation();
  }

  private async loadWebInformation() {
    await sp.web.lists.ensure(this._pnpDemoList01Name);

    const web = await sp.web.get();
    const subsites = await sp.web.webs.get();
    const fields = await sp.web.fields
      .select('InternalName')
      .filter('Hidden eq false and ReadOnlyField eq false and CanBeDeleted eq true')
      .get();

    this.setState({
      loading: false,
      webInformation: web,
      subsites: subsites,
      fields: fields
    });

    console.table(web); console.table(subsites); console.table(fields);
  }

  public render(): React.ReactElement<IPnPDemo01Props> {
    let { webInformation, subsites, loading, fields } = this.state;

    const loadingSpinner: JSX.Element = <Spinner size={SpinnerSize.medium} />;

    if (loading)
      return loadingSpinner;

    return (
      <div>
        <h1>{webInformation.Title} ({webInformation.WebTemplate})</h1>
        <p>{webInformation.Url} - {webInformation.Id}</p>
        <p>{`There ${subsites.length === 1 ? 'is' : 'are'} ${subsites.length} subsite${subsites.length === 1 ? '' : 's'} in total.`}</p>
        <p onClick={() => alert(fields.map(f => f.InternalName).join(', '))}><b>Fields:</b>${fields.length}</p>

        <h2>Batching</h2>
        <p><PrimaryButton
          text={`Create ${this._itemscount} items - ${this.state.itemsCreationPerformance || "..."}`}
          onClick={this._createItems}
          disabled={this.state.loadingCreatingItems}
        /></p>

        <p><PrimaryButton
          text={`Create ${this._itemscount} items (Batching) - ${this.state.itemsCreationBatchPerformance || "..."}`}
          onClick={this._createItemsUsingBatching}
          disabled={this.state.loadingCreatingItems}
        /></p>
      </div>
    );
  }

  private async _createItems() {
    this.setState({ ...this.state, loadingCreatingItems: true });
    const t0 = performance.now();

    const list = sp.web.lists.getByTitle(this._pnpDemoList01Name);
    for (let i = 0; i < this._itemscount; i++) {
      await list.items.add({}, "SP.Data.PnPDemo01ListItem");
    }

    const t1 = performance.now();
    const timeresult = (t1 - t0);
    this.setState({ ...this.state, itemsCreationPerformance: timeresult, loadingCreatingItems: false });
  }

  private async _createItemsUsingBatching() {
    this.setState({ ...this.state, loadingCreatingItems: true });
    const t0 = performance.now();

    const list = sp.web.lists.getByTitle(this._pnpDemoList01Name);
    const batch = sp.web.createBatch();
    for (let i = 0; i < this._itemscount; i++) {
      // Don't forget to pass the list entity type name
      // If we don't do this PnPjs will try to fetch it for us for all items
      list.items.inBatch(batch).add({}, "SP.Data.PnPDemo01ListItem");
    }
    await batch.execute();

    const t1 = performance.now();
    const timeresult = (t1 - t0);
    this.setState({ ...this.state, itemsCreationBatchPerformance: timeresult, loadingCreatingItems: false });
  }
}
