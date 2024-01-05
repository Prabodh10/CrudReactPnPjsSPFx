// import React, { useState, useEffect } from 'react';
import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './PnPjsExample.module.scss';
import { IPnPjsExampleProps } from './IPnPjsExampleProps';

// import interfaces
import { IFile, IResponseItem } from "./interfaces";

import { Caching } from "@pnp/queryable";
import { getSP } from "../pnpjsConfig";
import { SPFI, spfi } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";
import { Label, PrimaryButton } from 'office-ui-fabric-react';

// import { Label, PrimaryButton } from '@microsoft/office-ui-fabric-react-bundle';

interface IAsyncAwaitPnPJsProps {
  description: string;
}

// interface IIPnPjsExampleState {
//   items: IFile[];
//   errors: string[];
// }

const PnPjsExample: React.FC<IPnPjsExampleProps & IAsyncAwaitPnPJsProps> = (props: IPnPjsExampleProps) => {
  const LOG_SOURCE = "🅿PnPjsExample";
  const LIBRARY_NAME = "EmployeeDetails";
  const [items, setItems] = useState<IFile[]>([]);
 // const [errors, setErrors] = useState<string[]>([]);
  const _sp: SPFI = getSP();

  useEffect(() => {
    // componentDidMount logic
    _readAllFilesSize();

    // cleanup logic (equivalent to componentWillUnmount in class component)
    return () => {
      // cleanup code
    };
  }, []); // empty dependency array means this effect runs once after the first render

  const _readAllFilesSize = async (): Promise<void> => {
    try {
      const spCache = spfi(_sp).using(Caching({ store: "session" }));

      const response: IResponseItem[] = await spCache.web.lists
        .getByTitle(LIBRARY_NAME)
        .items
        .select("Id", "Title", "FileLeafRef", "File/Length")
        .expand("File/Length")();

      const newItems: IFile[] = response.map((item: IResponseItem) => ({
        Id: item.Id,
        Title: item.Title || "Unknown",
      }));

      setItems(newItems);
    } catch (err) {
      Logger.write(`${LOG_SOURCE} (_readAllFilesSize) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  };

  const _updateTitles = async (itemId: number): Promise<void> => {
    try {
      if (itemId !== undefined) {
        const spCache = spfi(_sp).using(Caching({ store: "session" }));
        const item = await spCache.web.lists
          .getByTitle(LIBRARY_NAME)
          .items.getById(itemId)
          .select("Id", "Title", "FileLeafRef", "File/Length")
          .expand("File/Length")();

        console.log(`Item Title: ${item.Title}`);

        const newTitle = await _promptForUserInput("Enter the new Title");

        if (newTitle !== undefined) {
          await spCache.web.lists
            .getByTitle(LIBRARY_NAME)
            .items.getById(itemId)
            .update({ Title: newTitle });

          const updatedItems = items.map(existingItem =>
            existingItem.Id === itemId ? { ...existingItem, Title: newTitle } : existingItem
          );

          setItems(updatedItems);
        }
      }
    } catch (err) {
      Logger.write(`${LOG_SOURCE} (_updateTitles) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  };

  const _createNewItem = async (): Promise<void> => {
    try {
      const employeeName = await _promptForUserInput("Enter the name of the new employee");

      if (employeeName !== undefined) {
        const newItem = await _sp.web.lists.getByTitle("EmployeeDetails").items.add({
          Title: employeeName,
        });

        const newItems = [...items, {
          Id: newItem.data.Id,
          Title: employeeName,
          Size: 0,
          Name: "",
        }];

        setItems(newItems);
      }
    } catch (err) {
      Logger.write(`${LOG_SOURCE} (_createNewItem) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  };

  const _promptForUserInput = async (prompt: string): Promise<string | undefined> => {
    return new Promise(resolve => {
      const result = window.prompt(prompt);
      resolve(result !== null ? result : undefined);
    });
  };

  const _deleteItem = async (itemId: number): Promise<void> => {
    try {
      await _sp.web.lists.getByTitle("EmployeeDetails").items.getById(itemId).delete();

      const updatedItems = items.filter(item => item.Id !== itemId);
      setItems(updatedItems);
    } catch (err) {
      Logger.write(`${LOG_SOURCE} (_deleteItem) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  };

  return (
    <div className={`${styles.pnPjsExample} ${styles.centerTable}`}>
      <Label>Welcome to PnP JS CRUD React</Label>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <Label>List of items:</Label>
        <PrimaryButton onClick={_createNewItem} style={{ marginLeft: 'auto' }}>Create New Item</PrimaryButton>
      </div>
      <table width="100%">
        <thead>
          <tr>
            <th>ID</th>
            <th>Title</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {items.map((item, idx) => (
            <tr key={idx}>
              <td>{item.Id}</td>
              <td>{item.Title}</td>
              <td>
                <PrimaryButton onClick={() => _updateTitles(item.Id)}>Edit</PrimaryButton>&nbsp;&nbsp;&nbsp;
                <PrimaryButton onClick={() => _deleteItem(item.Id)}>Delete</PrimaryButton>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default PnPjsExample;
