import * as React from 'react';
import { useState, useEffect } from 'react';
import { BaseService, ISiteInfo, IListCheckResult } from '../../../services/BaseService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { 
  PrimaryButton, 
  DefaultButton,
  MessageBar, 
  MessageBarType, 
  Spinner, 
  SpinnerSize,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  TextField,
  Stack,
  StackItem
} from '@fluentui/react';

export interface IConnectionTestProps {
  context: WebPartContext;
}

// Интерфейс для элементов списка в DetailsList
interface IListItem {
  name: string;
  status: string;
  itemCount: string | number;
  details: string;
}

export const ConnectionTest: React.FC<IConnectionTestProps> = (props) => {
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [siteInfo, setSiteInfo] = useState<ISiteInfo | null>(null);
  const [listsInfo, setListsInfo] = useState<IListCheckResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [customListName, setCustomListName] = useState<string>('');
  const [prevSiteUrl, setPrevSiteUrl] = useState<string>('');
  const [baseService, setBaseService] = useState<BaseService | null>(null);

  useEffect(() => {
    const service = new BaseService(props.context, "ConnectionTest");
    setBaseService(service);
    
    // Безопасно получаем URL предыдущего сайта из BaseService
    // Используем явное приведение типа вместо @ts-ignore
    setPrevSiteUrl((service as any)._prevSiteUrl || '');
  }, [props.context]);

  // Тестирование подключения к сайту
  const handleTestConnection = async (): Promise<void> => {
    if (!baseService) return;

    setIsLoading(true);
    setSiteInfo(null);
    setError(null);

    try {
      const webInfo = await baseService.testPrevSiteConnection();
      setSiteInfo(webInfo);
    } catch (err) {
      console.error("Connection test failed:", err);
      setError(`Connection test failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Проверка доступности списков
  const handleCheckLists = async (): Promise<void> => {
    if (!baseService) return;

    setIsLoading(true);
    setListsInfo(null);
    setError(null);

    try {
      const results = await baseService.checkAllRequiredLists();
      setListsInfo(results);
    } catch (err) {
      console.error("List check failed:", err);
      setError(`List check failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Проверка пользовательского списка
  const handleCheckCustomList = async (): Promise<void> => {
    if (!baseService) return;
    
    if (!customListName.trim()) {
      setError("Please enter a list name");
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      const listInfo = await baseService.checkListExists(customListName);
      
      setListsInfo({ [customListName]: listInfo });
    } catch (err) {
      console.error(`List check failed for "${customListName}":`, err);
      setError(`List check failed for "${customListName}": ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Колонки для таблицы списков
  const columns: IColumn[] = [
    {
      key: 'listName',
      name: 'List Name',
      fieldName: 'name',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'status',
      minWidth: 100,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IListItem) => (
        <span style={{ 
          color: item.status === 'OK' ? 'green' : 'red', 
          fontWeight: 'bold' 
        }}>
          {item.status}
        </span>
      )
    },
    {
      key: 'itemCount',
      name: 'Item Count',
      fieldName: 'itemCount',
      minWidth: 100,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'details',
      name: 'Details',
      fieldName: 'details',
      minWidth: 200,
      isResizable: true
    }
  ];

  // Преобразуем информацию о списках в формат для DetailsList
  const getListItems = (): IListItem[] => {
    if (!listsInfo) return [];

    return Object.keys(listsInfo).map(listName => {
      const info = listsInfo[listName];
      
      if ('error' in info) {
        return {
          name: listName,
          status: 'Error',
          itemCount: '-',
          details: info.error
        };
      }
      
      return {
        name: listName,
        status: 'OK',
        itemCount: info.ItemCount || 0,
        details: `ID: ${info.Id}`
      };
    });
  };

  return (
    <div style={{ padding: '20px' }}>
      <Stack tokens={{ childrenGap: 15 }}>
        <h2>Previous Site Connection Test</h2>
        
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <StackItem>
            <TextField
              label="Previous Site URL"
              value={prevSiteUrl}
              readOnly
              styles={{ field: { width: 400 } }}
            />
          </StackItem>
        </Stack>
        
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <StackItem>
            <PrimaryButton 
              text="Test Connection" 
              onClick={handleTestConnection}
              disabled={isLoading}
            />
          </StackItem>
          
          <StackItem>
            <DefaultButton 
              text="Check Standard Lists" 
              onClick={handleCheckLists}
              disabled={isLoading}
            />
          </StackItem>
        </Stack>

        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 10 }}>
          <StackItem grow>
            <TextField 
              label="Custom List Name" 
              value={customListName}
              onChange={(_, newValue) => setCustomListName(newValue || '')}
              disabled={isLoading}
              styles={{ root: { width: 200 } }}
            />
          </StackItem>
          
          <StackItem>
            <DefaultButton 
              text="Check Custom List" 
              onClick={handleCheckCustomList}
              disabled={isLoading || !customListName.trim()}
            />
          </StackItem>
        </Stack>

        {isLoading && (
          <Stack>
            <Spinner size={SpinnerSize.medium} label="Testing connection..." />
          </Stack>
        )}

        {error && (
          <Stack>
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          </Stack>
        )}

        {siteInfo && (
          <Stack tokens={{ padding: '10px 0' }}>
            <MessageBar messageBarType={MessageBarType.success}>
              Successfully connected to: {siteInfo.Title}
            </MessageBar>
            <div style={{ marginTop: '10px' }}>
              <strong>Site URL:</strong> {siteInfo.Url}<br />
              <strong>Site ID:</strong> {siteInfo.Id}<br />
              <strong>Created:</strong> {new Date(siteInfo.Created).toLocaleString()}<br />
              <strong>Last Modified:</strong> {new Date(siteInfo.LastItemModifiedDate).toLocaleString()}
            </div>
          </Stack>
        )}

        {listsInfo && (
          <Stack>
            <h3>List Check Results</h3>
            <DetailsList
              items={getListItems()}
              columns={columns}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.none}
              isHeaderVisible={true}
            />
          </Stack>
        )}
      </Stack>
    </div>
  );
};