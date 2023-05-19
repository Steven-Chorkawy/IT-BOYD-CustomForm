import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import styles from './Byod.module.scss';

export interface IByodProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'Byod';

export default class Byod extends React.Component<IByodProps, {}> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: Byod mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: Byod unmounted');
  }

  public render(): React.ReactElement<{}> {
    return <div className={styles.byod} />;
  }
}
