import * as React from "react";
import styles from "./Speiseplan.module.scss";
import type { ISpeiseplanProps } from "./ISpeiseplanProps";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

export const Speiseplan: React.FunctionComponent<ISpeiseplanProps> = ({
  hasTeamsContext,
  isDarkTheme,
  userDisplayName,
  environmentMessage,
  description,
  context,
}: ISpeiseplanProps) => {
  const [lists, setLists] = React.useState<{ [key: string]: Array<any> }>({});
  const [loading, setLoading] = React.useState(true);
  const sp = spfi("https://sfzbbw.sharepoint.com").using(SPFx(context));

  const fields = ["Datum", "Gericht 1", "Gericht 2", "Suppe", "Salat"];

  React.useEffect(() => {
    const fetchData = async () => {
      try {
        const fetchedLists: { [key: string]: Array<any> } = {};
        for (const field of fields) {
          const actualName = await sp.web.lists.getByTitle("Speiseplan").fields.getByTitle(field)();
          const list = await sp.web.lists
            .getByTitle("Speiseplan")
            .items.select(actualName.EntityPropertyName)
            .top(50)
            .orderBy("Datum", false)();
          fetchedLists[actualName.EntityPropertyName] = list;
        }
        setLists(fetchedLists);
      } catch (error) {
        console.error("Error fetching Speiseplan list items", error);
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, [context]);

  const renderList = (field: string) => {
    const fieldName = Object.keys(lists).filter(function (name) {
      return name.indexOf(field) !== -1;
    })[0];
    if (!fieldName || lists[fieldName].length === 0) {
      return <div>No Items found</div>;
    }

    return (
      <ul>
        {lists[fieldName].map(item => (
          <li key={item.Id}>{item[fieldName]}</li>
        ))}
      </ul>
    );
  };

  return (
    <section className={`${styles.speiseplan} ${hasTeamsContext ? styles.teams : ""}`}>
      <div className={styles.welcome}>
        <div>
          <strong>3 Speiseplan</strong>
        </div>
        {loading ? <div>Loading...</div> : renderList("Datum")}
      </div>
      <div>{loading ? <div>Loading...</div> : renderList("Gericht 1")}</div>
      <div>{loading ? <div>Loading...</div> : renderList("Gericht 2")}</div>
      <div>{loading ? <div>Loading...</div> : renderList("Suppe")}</div>
      <div>{loading ? <div>Loading...</div> : renderList("Salat")}</div>
    </section>
  );
};

export default Speiseplan;
