import * as React from 'react';
import styles from './ConnectionTest.module.scss';
import { BaseService, ISiteInfo, IListCheckResult, IListInfo, IListItem } from '../../../services/BaseService';
import { IConnectionTestProps } from './IConnectionTestProps';
import { IConnectionTestState } from './IConnectionTestState';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, DefaultButton, TextField, Stack, StackItem, Label, Spinner, SpinnerSize, MessageBar, MessageBarType } from 'office-ui-fabric-react';

export default class ConnectionTest extends React.Component<IConnectionTestProps, IConnectionTestState> {
  private _baseService: BaseService;
  
  constructor(props: IConnectionTestProps) {
    super(props);
    // Инициализация сервиса
    this._baseService = new BaseService(this.props.context, "ConnectionTest");
    
    // Инициализация состояния
    this.state = {
      loading: false,
      error: null,
      prevSiteUrl: this._baseService.getPrevSiteUrl(),
      siteInfo: null,
      listsCheckResult: null,
      updateStatus: null,
      updateTitle: "",
      listTitle: "Staff",
      updatedItem: null
    };
  }

  public render(): React.ReactElement<IConnectionTestProps> {
    const { loading, error, prevSiteUrl, siteInfo, listsCheckResult, updateStatus, updatedItem } = this.state;
    
    return (
      <div className={styles.connectionTest}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <h2>Тестирование подключения к удаленному сайту</h2>
              <p>Этот инструмент проверяет возможность подключения к удаленному сайту SharePoint через Graph API.</p>
              
              {/* URL удаленного сайта */}
              <Label>URL удаленного сайта:</Label>
              <TextField 
                readOnly
                value={prevSiteUrl}
                className={styles.textField}
              />
              
              <Stack horizontal tokens={{ childrenGap: 10 }} className={styles.buttonContainer}>
                {/* Кнопка проверки подключения */}
                <PrimaryButton 
                  text="Проверить подключение"
                  onClick={this._testConnection}
                  disabled={loading}
                />
                
                {/* Кнопка проверки списков */}
                <DefaultButton 
                  text="Проверить списки"
                  onClick={this._checkLists}
                  disabled={loading || !siteInfo}
                />
              </Stack>
              
              {/* Секция обновления элемента списка */}
              <div className={styles.updateSection}>
                <h3>Обновление элемента списка</h3>
                <Stack tokens={{ childrenGap: 10 }}>
                  <StackItem>
                    <Label>Название списка:</Label>
                    <TextField 
                      value={this.state.listTitle}
                      onChange={this._onListTitleChange}
                      className={styles.textField}
                    />
                  </StackItem>
                  <StackItem>
                    <Label>Новое значение Title:</Label>
                    <TextField 
                      value={this.state.updateTitle}
                      onChange={this._onUpdateTitleChange}
                      className={styles.textField}
                    />
                  </StackItem>
                  <StackItem>
                    <PrimaryButton 
                      text="Обновить элемент"
                      onClick={this._updateItem}
                      disabled={loading || !siteInfo || !this.state.updateTitle}
                    />
                  </StackItem>
                </Stack>
              </div>
              
              {/* Отображение индикатора загрузки */}
              {loading && (
                <div className={styles.spinner}>
                  <Spinner size={SpinnerSize.large} label="Загрузка..." />
                </div>
              )}
              
              {/* Отображение ошибок */}
              {error && (
                <MessageBar messageBarType={MessageBarType.error} className={styles.messageBar}>
                  {error}
                </MessageBar>
              )}
              
              {/* Отображение информации о сайте */}
              {siteInfo && !loading && (
                <div className={styles.infoSection}>
                  <h3>Информация о сайте</h3>
                  <div className={styles.infoItem}>
                    <strong>Название:</strong> {siteInfo.Title}
                  </div>
                  <div className={styles.infoItem}>
                    <strong>URL:</strong> {siteInfo.Url}
                  </div>
                  <div className={styles.infoItem}>
                    <strong>Дата создания:</strong> {new Date(siteInfo.Created).toLocaleString()}
                  </div>
                  <div className={styles.infoItem}>
                    <strong>Последнее изменение:</strong> {new Date(siteInfo.LastItemModifiedDate).toLocaleString()}
                  </div>
                  {siteInfo.Description && (
                    <div className={styles.infoItem}>
                      <strong>Описание:</strong> {siteInfo.Description}
                    </div>
                  )}
                </div>
              )}
              
              {/* Отображение результатов проверки списков */}
              {listsCheckResult && !loading && (
                <div className={styles.infoSection}>
                  <h3>Проверка списков</h3>
                  {Object.keys(listsCheckResult).map(listName => {
                    const result = listsCheckResult[listName];
                    return (
                      <div key={listName} className={styles.listInfo}>
                        <h4>{listName}</h4>
                        {'error' in result ? (
                          <MessageBar messageBarType={MessageBarType.error}>
                            {result.error}
                          </MessageBar>
                        ) : (
                          <div>
                            <div className={styles.infoItem}>
                              <strong>ID:</strong> {result.Id}
                            </div>
                            <div className={styles.infoItem}>
                              <strong>Количество элементов:</strong> {result.ItemCount}
                            </div>
                            {result.Description && (
                              <div className={styles.infoItem}>
                                <strong>Описание:</strong> {result.Description}
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
              
              {/* Отображение статуса обновления */}
              {updateStatus && (
                <MessageBar 
                  messageBarType={updateStatus.type} 
                  className={styles.messageBar}
                >
                  {updateStatus.message}
                </MessageBar>
              )}
              
              {/* Отображение обновленного элемента */}
              {updatedItem && !loading && (
                <div className={styles.infoSection}>
                  <h3>Обновленный элемент</h3>
                  <div className={styles.infoItem}>
                    <strong>ID:</strong> {updatedItem.Id}
                  </div>
                  <div className={styles.infoItem}>
                    <strong>Title:</strong> {updatedItem.Title}
                  </div>
                  {Object.keys(updatedItem)
                    .filter(key => key !== 'Id' && key !== 'Title' && typeof updatedItem[key] !== 'object')
                    .map(key => (
                      <div key={key} className={styles.infoItem}>
                        <strong>{key}:</strong> {String(updatedItem[key])}
                      </div>
                    ))}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }

  /**
   * Обработчик изменения названия списка
   */
  private _onListTitleChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ listTitle: newValue || "" });
  }

  /**
   * Обработчик изменения значения для обновления
   */
  private _onUpdateTitleChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ updateTitle: newValue || "" });
  }

  /**
   * Проверяет подключение к удаленному сайту
   */
  private _testConnection = async (): Promise<void> => {
    this.setState({ 
      loading: true, 
      error: null, 
      siteInfo: null,
      listsCheckResult: null,
      updateStatus: null,
      updatedItem: null
    });
    
    try {
      const siteInfo = await this._baseService.testPrevSiteConnection();
      this.setState({ 
        siteInfo,
        loading: false 
      });
    } catch (error) {
      this.setState({ 
        error: `Ошибка подключения: ${error instanceof Error ? error.message : String(error)}`,
        loading: false 
      });
    }
  }

  /**
   * Проверяет доступные списки на удаленном сайте
   */
  private _checkLists = async (): Promise<void> => {
    this.setState({ 
      loading: true, 
      error: null,
      listsCheckResult: null,
      updateStatus: null,
      updatedItem: null
    });
    
    try {
      const listsCheckResult = await this._baseService.checkAllRequiredLists();
      this.setState({ 
        listsCheckResult,
        loading: false 
      });
    } catch (error) {
      this.setState({ 
        error: `Ошибка проверки списков: ${error instanceof Error ? error.message : String(error)}`,
        loading: false 
      });
    }
  }

  /**
   * Обновляет первый элемент в выбранном списке
   */
  private _updateItem = async (): Promise<void> => {
    const { listTitle, updateTitle } = this.state;
    
    if (!updateTitle) {
      this.setState({ 
        updateStatus: {
          type: MessageBarType.warning,
          message: "Введите значение для обновления"
        }
      });
      return;
    }
    
    this.setState({ 
      loading: true,
      error: null,
      updateStatus: null,
      updatedItem: null
    });
    
    try {
      const updatedItem = await this._baseService.updateFirstItemTitleField(listTitle, updateTitle);
      
      this.setState({ 
        updatedItem,
        updateStatus: {
          type: MessageBarType.success,
          message: `Элемент успешно обновлен в списке "${listTitle}"`
        },
        loading: false,
        updateTitle: "" // Очищаем поле после успешного обновления
      });
    } catch (error) {
      this.setState({ 
        updateStatus: {
          type: MessageBarType.error,
          message: `Ошибка обновления элемента: ${error instanceof Error ? error.message : String(error)}`
        },
        loading: false 
      });
    }
  }
}