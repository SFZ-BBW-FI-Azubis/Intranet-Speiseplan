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
  const [lists, setLists] = React.useState<{ [key: string]: Array<any> }>({
    list0: [],
    list1: [],
    list2: [],
    list3: [],
    list4: [],
  });
  const [fieldNames, setFieldNames] = React.useState<{ [key: string]: string }>(
    {}
  );
  const sp = spfi("https://sfzbbw.sharepoint.com").using(SPFx(context));

  const fetchListData = async (fieldName: string, listKey: string) => {
    try {
      const field = await sp.web.lists
        .getByTitle("Speiseplan")
        .fields.getByTitle(fieldName)();
      const items = await sp.web.lists
        .getByTitle("Speiseplan")
        .items.select(field.EntityPropertyName)
        .top(10)
        .orderBy("Datum", false)();

      if (listKey === "list0") {
        items.forEach((item) => {
          item[field.EntityPropertyName] = item[field.EntityPropertyName].slice(
            5,
            -10
          );
        });
      }

      setLists((prevLists) => ({ ...prevLists, [listKey]: items }));
      setFieldNames((prevFieldNames) => ({
        ...prevFieldNames,
        [listKey]: field.EntityPropertyName,
      }));
    } catch (error) {
      console.error(`Error fetching ${fieldName} list items`, error);
    }
  };

  React.useEffect(() => {
    const listFields = ["Datum", "Gericht 1", "Gericht 2", "Suppe", "Salat"];
    const fetchAllLists = async () => {
      await Promise.all(
        listFields.map((field, index) => fetchListData(field, `list${index}`))
      );
    };
    fetchAllLists();
  }, [context]);

  const renderList = (listKey: string) => {
    const list = lists[listKey];
    const fieldName = fieldNames[listKey];
    return list.length > 0 ? (
      <ul>
        {list.map((item) => (
          <li key={item.Id}>{item[fieldName]}</li>
        ))}
      </ul>
    ) : (
      <div>No Items found</div>
    );
  };

  return (
    <section
      className={`${styles.speiseplan} ${hasTeamsContext ? styles.teams : ""}`}
    >
      <div>
        <h1>
          <strong>Speiseplan</strong>
        </h1>
        <div className={styles.listHolder}>
          <div className={styles.listDate}>
            <h1>Datum</h1>
            <div>{renderList("list0")}</div>
          </div>
          <div className={styles.listMeal1}>
            <h1>Speise 1</h1>
            <div>{renderList("list1")}</div>
          </div>
          <div className={styles.listMeal2}>
            <h1>Speise 2</h1>
            <div>{renderList("list2")}</div>
          </div>
          <div className={styles.listSoup}>
            <h1>Suppe</h1>
            <div>{renderList("list3")}</div>
          </div>
          <div className={styles.listSalad}>
            <h1>Salat</h1>
            <div>{renderList("list4")}</div>
          </div>
        </div>
      </div>
    </section>
  );
};

export default Speiseplan;
