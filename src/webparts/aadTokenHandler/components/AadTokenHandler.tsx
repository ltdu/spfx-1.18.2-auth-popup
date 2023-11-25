import * as React from 'react';
import styles from './AadTokenHandler.module.scss';
import type { IAadTokenHandlerProps } from './IAadTokenHandlerProps';
import { DefaultButton } from '@fluentui/react/lib/Button';

export default class AadTokenHandler extends React.Component<IAadTokenHandlerProps> {
  public render(): React.ReactElement<IAadTokenHandlerProps> {
    const {
      context,
      redirectionRequired,
      redirectionUrl,
      popupRequired,
      popup,
      invoke,
      log
    } = this.props;

    const rows = log.map(entry => (
      <>
        <p>
          <pre>{entry}</pre>
        </p>
        <hr />
      </>
    ));

    const canInvoke = true;

    return (
      <section className={styles.aadTokenHandler}>
        <div className={styles.welcome}>
          <h2>SPFX: 1.18.2 | Web part version: {context.manifest.version}</h2>
        </div>
        <div>
          {canInvoke &&
            <DefaultButton className={styles.button} onClick={invoke}>{"Invoke service with AAD authentication"}</DefaultButton>
          }
          {redirectionRequired &&
            <DefaultButton className={styles.button} href={redirectionUrl}>{"Refresh to authenticate"}</DefaultButton>
          }
          {popupRequired &&
            <DefaultButton className={styles.button} onClick={popup}>{"Open pop up to authenticate"}</DefaultButton>
          }
        </div>
        <div>
          {rows}
        </div>
      </section>
    );
  }
}
