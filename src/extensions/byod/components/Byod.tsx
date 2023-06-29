import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { WidgetSize, Dashboard } from '@pnp/spfx-controls-react/lib/Dashboard';

export interface IByodProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'Byod';

export default class Byod extends React.Component<IByodProps, {}> {
  /**
   *
   */
  constructor(props: IByodProps) {
    super(props);
    this._item = this.props.context.item;
    console.log(this.props);
    console.log('Item:');
    console.log(this.props.context.item);
  }

  private _item: any;

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: Byod mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: Byod unmounted');
  }

  public render(): React.ReactElement<{}> {
    return <div>
      <Dashboard
        widgets={[{
          title: "BYOD Submission",
          // widgetActionGroup: calloutItemsExample,
          size: WidgetSize.Box,
          body: [
            {
              id: "t1",
              title: "Tab 1",
              content: (
                <div>
                  <h1>{this._item.Title}</h1>
                  <p>Submission Date: {this._item.Date1 && new Date(this._item.Date1).toLocaleDateString()}</p>
                  <p>Final Approval Date: {this._item.FinalApprovalDate && new Date(this._item.FinalApprovalDate).toLocaleDateString()}</p>
                  <p>Device/ Model: {this._item.DeviceManufacturer}/ {this._item.DeviceModel}</p>
                  <p>Operating System: {this._item.DeviceOperatingSystem}</p>
                  <p>Wireless Provider: {this._item.WirelessProvider}</p>
                  <p>Contract Type: {this._item.ContractType}</p>
                  <p>Contract End Date: {this._item.ContractEndDate && new Date(this._item.ContractEndDate).toLocaleDateString('en-US')}</p>
                </div>
              ),
            },
          ]
        },
        {
          title: "Approval Status",
          size: WidgetSize.Box,
          body: [
            {
              id: 'c2t1',
              title: 'Card 2 Title',
              content: (
                <div>
                  <h1>{this._item.OData__Status}</h1>
                  <hr />
                  <h4>Approval Comments</h4>
                  <pre>{this._item.ApprovalComments}</pre>
                  <hr />
                  <h4>Approval Summary</h4>
                  <pre>{this._item.ApprovalSummary}</pre>
                </div>
              )
            }
          ]
        },
          // {
          //   title: "Card 3",
          //   size: WidgetSize.Box,
          //   //link: linkExample,
          //   body: [
          //     {
          //       id: 'c2t1',
          //       title: 'Card 3 Title - JSON data',
          //       content: (<div>
          //         {JSON.stringify(this._item)}
          //       </div>)
          //     }
          //   ]
          // }
        ]} />
    </div>;
  }
}
