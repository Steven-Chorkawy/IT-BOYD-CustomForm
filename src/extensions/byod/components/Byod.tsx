import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { WidgetSize, Dashboard } from '@pnp/spfx-controls-react/lib/Dashboard';

import { Icon } from '@fluentui/react';


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
  }

  private _item: any;

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: Byod mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: Byod unmounted');
  }

  public render(): React.ReactElement<{}> {
    const linkExample = { href: "#" };
    const calloutItemsExample = [
      {
        id: "action_1",
        title: "Info",
        icon: <Icon iconName={'Edit'} />,
      },
      { id: "action_2", title: "Popup", icon: <Icon iconName={'Add'} /> },
    ];

    return <div>
      <Dashboard
        widgets={[{
          title: "BYOD Submission",
          widgetActionGroup: calloutItemsExample,
          size: WidgetSize.Box,
          body: [
            {
              id: "t1",
              title: "Tab 1",
              content: (
                <div>
                  <h1>{this._item.Title}</h1>
                  <h4>Submission Date: {this._item.Date}</h4>
                  <p>Device Model: {this._item.DeviceModel}</p>
                  <p>Wireless Provider: {this._item.WirelessProvider}</p>
                  <p>Contract Type: {this._item.ContractType}</p>
                </div>
              ),
            },
          ],
          link: linkExample,
        },
        {
          title: "Approval Status",
          size: WidgetSize.Double,
          link: linkExample,
          body: [
            {
              id: 'c2t1',
              title: 'Card 2 Title',
              content: (
                <div>
                  <h2>Status: {this._item.Status}</h2>
                  <p>Manager Approval: ...name here, date here, comments here...</p>
                  <p>IT Approval: ...name here, date here, comments here..</p>
                </div>
              )
            }
          ]
        },
        {
          title: "Card 3",
          size: WidgetSize.Double,
          link: linkExample,
          body: [
            {
              id: 'c2t1',
              title: 'Card 3 Title - JSON data',
              content: (<div>
                {JSON.stringify(this._item)}
              </div>)
            }
          ]
        }]} />
    </div>;
  }
}
