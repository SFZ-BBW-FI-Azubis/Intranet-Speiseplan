import * as React from 'react';
import styles from './Speiseplan.module.scss';
import type { ISpeiseplanProps } from './ISpeiseplanProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";

export const Speiseplan: React.FunctionComponent<ISpeiseplanProps> =  ({ hasTeamsContext, isDarkTheme, userDisplayName, environmentMessage, description, context }: ISpeiseplanProps) => {
  const [list, setList] = React.useState<Array<any>>([])
  const sp = spfi().using(SPFx(context))

  React.useEffect(() => {
    (async () => {
      const temp:any = await sp.web.lists.getByTitle("Speiseplan Test")()
      console.log("test",temp)
      setList(temp)
      list
    })()
  }, [])

  return (
    <section className={`${styles.speiseplan} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
      </div>
    </section>
  );
}

export default Speiseplan;