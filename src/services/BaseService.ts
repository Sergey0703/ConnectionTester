// src/services/BaseService.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/lists/web";
import { MSGraphClientV3 } from '@microsoft/sp-http';

// Интерфейс для информации о сайте
export interface ISiteInfo {
  Id: string;
  Title: string;
  Url: string;
  Created: string;
  LastItemModifiedDate: string;
  Description?: string;
  ServerRelativeUrl?: string;
  WebTemplate?: string;
  [key: string]: unknown; // Индексная сигнатура для дополнительных полей
}

// Интерфейс для информации о списке
export interface IListInfo {
  Id: string;
  Title: string;
  ItemCount: number;
  Description?: string;
  DefaultViewUrl?: string;
  LastItemModifiedDate?: string;
  [key: string]: unknown; // Индексная сигнатура для дополнительных полей
}

// Интерфейс для результатов проверки списков
export interface IListCheckResult {
  [listName: string]: IListInfo | { error: string };
}

// Интерфейс для обновленного элемента списка
export interface IListItem {
  Id: number;
  Title: string;
  [key: string]: unknown;
}

export class BaseService {
  protected _sp: SPFI;
  protected _logSource: string;
  protected _context: WebPartContext;
  
  // URL предыдущего сайта
  protected _prevSiteUrl: string = "https://kpfaie.sharepoint.com/sites/KPFAData";
  
  // ID сайта для Graph API (заполняется при инициализации)
  protected _targetSiteId: string | null = null;
  protected _targetSiteDriveId: string | null = null;
  
  // Флаг авторизации
  protected _isAuthorized: boolean = false;
  
  constructor(context: WebPartContext, logSource: string) {
    this._context = context;
    // Инициализируем PnP JS с контекстом для текущего сайта
    this._sp = spfi().using(SPFx(context));
    this._logSource = logSource;
    
    // Инициализируем Graph авторизацию и получаем ID удаленного сайта
    this.initGraphAuthorization();
  }

  /**
   * Инициализирует авторизацию через Graph API и получает ID сайта
   */
  private async initGraphAuthorization(): Promise<void> {
    try {
      this.logInfo("Initializing Graph client and authorizing access to remote site...");
      
      // Получаем Graph клиент, который автоматически включает токен авторизации
      const graphClient: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
      this.logInfo("Successfully obtained Graph client with authorization token");
      
      // Извлекаем домен и относительный путь из URL
      const url = new URL(this._prevSiteUrl);
      const hostname = url.hostname;
      const pathname = url.pathname;
      
      // Проверяем доступ к удаленному сайту, используя авторизационный токен,
      // предоставленный приложению (не пользователю)
      this.logInfo(`Verifying authorized access to remote site at ${hostname}${pathname}`);
      
      try {
        // Запрашиваем ID сайта по хостнейму и пути
        const response = await graphClient
          .api(`/sites/${hostname}:${pathname}`)
          .get();
        
        this._targetSiteId = response.id;
        this._targetSiteDriveId = response.drive?.id;
        this._isAuthorized = true;
        
        this.logInfo(`Authorization successful. Site ID: ${this._targetSiteId}`);
      } catch (authError) {
        this._isAuthorized = false;
        
        if (authError.statusCode === 401 || authError.statusCode === 403) {
          this.logError("Authorization to remote site failed - insufficient permissions.");
          this.logError("Ensure that app permissions are approved in SharePoint Admin Center:");
          this.logError("1. Go to SharePoint Admin Center > Advanced > API access");
          this.logError("2. Approve pending requests for Microsoft Graph permissions");
        } else {
          this.logError(`Remote site authorization error: ${authError.message}`);
        }
        
        throw new Error(`Failed to authorize access to remote site: ${authError.message}`);
      }
    } catch (error) {
      this.logError(`Failed to initialize Graph authorization: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Убеждается, что авторизация на удаленный сайт выполнена
   * @returns Promise, который разрешается, если авторизация успешна
   */
  private async ensureAuthorization(): Promise<void> {
    if (!this._isAuthorized || !this._targetSiteId) {
      this.logInfo("Re-initializing Graph authorization...");
      await this.initGraphAuthorization();
      
      if (!this._isAuthorized || !this._targetSiteId) {
        throw new Error("Authorization to remote site failed. Check application permissions.");
      }
    }
  }

  /**
   * Получает URL предыдущего сайта
   * @returns URL предыдущего сайта
   */
  public getPrevSiteUrl(): string {
    return this._prevSiteUrl;
  }

  /**
   * Проверяет авторизацию и соединение с предыдущим сайтом
   * @returns Promise с информацией о веб-сайте
   */
  public async testPrevSiteConnection(): Promise<ISiteInfo> {
    try {
      // Убедимся, что у нас есть авторизация на удаленный сайт
      await this.ensureAuthorization();
      
      // Получаем информацию о сайте через Graph API с авторизационным токеном
      const graphClient: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
      const siteData = await graphClient
        .api(`/sites/${this._targetSiteId}`)
        .get();
      
      this.logInfo(`Successfully connected to previous site: ${siteData.displayName}`);
      
      // Преобразуем данные из Graph API в наш интерфейс ISiteInfo
      const siteInfo: ISiteInfo = {
        Id: siteData.id,
        Title: siteData.displayName,
        Url: siteData.webUrl,
        Created: siteData.createdDateTime,
        LastItemModifiedDate: siteData.lastModifiedDateTime
      };
      
      // Добавляем дополнительные свойства
      if (siteData.description) siteInfo.Description = siteData.description;
      if (siteData.webUrl) siteInfo.ServerRelativeUrl = new URL(siteData.webUrl).pathname;
      if (siteData.template) siteInfo.WebTemplate = siteData.template.displayName;
      
      return siteInfo;
    } catch (error) {
      this.logError(`Failed to connect to previous site: ${error instanceof Error ? error.message : String(error)}`);
      throw error;
    }
  }

  /**
   * Проверяет доступность списка на предыдущем сайте используя авторизованный доступ
   * @param listTitle Название списка для проверки
   * @returns Promise с информацией о списке или ошибкой
   */
  public async checkListExists(listTitle: string): Promise<IListInfo> {
    try {
      // Убедимся, что у нас есть авторизация на удаленный сайт
      await this.ensureAuthorization();
      
      // Получаем Graph клиент с авторизационным токеном
      const graphClient: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
      
      // Получаем список с использованием авторизованного доступа
      // Используем фильтрацию на стороне сервера для оптимизации
      const listsResponse = await graphClient
        .api(`/sites/${this._targetSiteId}/lists`)
        .filter(`displayName eq '${listTitle}'`)
        .get();
      
      if (!listsResponse.value || listsResponse.value.length === 0) {
        throw new Error(`List "${listTitle}" not found`);
      }
      
      const listData = listsResponse.value[0];
      
      // Получаем количество элементов списка с авторизованным доступом
      const itemsResponse = await graphClient
        .api(`/sites/${this._targetSiteId}/lists/${listData.id}/items`)
        .count(true)
        .top(1)
        .get();
      
      const itemCount = itemsResponse["@odata.count"] || 0;
      
      this.logInfo(`Successfully accessed list "${listTitle}" with ${itemCount} items`);
      
      // Преобразуем данные из Graph API в наш интерфейс IListInfo
      const listInfo: IListInfo = {
        Id: listData.id,
        Title: listData.displayName,
        ItemCount: itemCount
      };
      
      // Добавляем дополнительные свойства
      if (listData.description) listInfo.Description = listData.description;
      if (listData.webUrl) listInfo.DefaultViewUrl = listData.webUrl;
      if (listData.lastModifiedDateTime) listInfo.LastItemModifiedDate = listData.lastModifiedDateTime;
      
      return listInfo;
    } catch (error) {
      this.logError(`Failed to access list "${listTitle}": ${error instanceof Error ? error.message : String(error)}`);
      throw error;
    }
  }

  /**
   * Проверяет все необходимые списки на предыдущем сайте используя авторизованный доступ
   * @returns Promise с результатами проверки
   */
  public async checkAllRequiredLists(): Promise<IListCheckResult> {
    const requiredLists = [
      "Staff",
      "StaffGroups",
      "GroupMembers",
      "WeeklySchedule",
      "TypeOfWorkers"
    ];
    
    const results: IListCheckResult = {};
    
    // Сначала проверяем авторизацию
    try {
      await this.ensureAuthorization();
    } catch (authError) {
      // Если авторизация не удалась, возвращаем ошибку для всех списков
      for (const listTitle of requiredLists) {
        results[listTitle] = {
          error: "Authorization to remote site failed. Check application permissions."
        };
      }
      return results;
    }
    
    // Если авторизация успешна, проверяем каждый список
    for (const listTitle of requiredLists) {
      try {
        results[listTitle] = await this.checkListExists(listTitle);
      } catch (error) {
        results[listTitle] = { 
          error: error instanceof Error ? error.message : String(error)
        };
      }
    }
    
    return results;
  }

  /**
   * Обновляет значение поля Title в первом элементе списка используя авторизованный доступ
   * @param listTitle Название списка
   * @param newValue Новое значение для поля Title
   * @returns Promise с обновленным элементом
   */
  public async updateFirstItemTitleField(listTitle: string, newValue: string): Promise<IListItem> {
    try {
      // Убедимся, что у нас есть авторизация на удаленный сайт
      await this.ensureAuthorization();
      
      // Получаем Graph клиент с авторизационным токеном
      const graphClient: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
      
      this.logInfo(`Using authorized access to update item in list "${listTitle}"`);
      
      // Шаг 1: Найти список по имени
      const listsResponse = await graphClient
        .api(`/sites/${this._targetSiteId}/lists`)
        .filter(`displayName eq '${listTitle}'`)
        .get();
      
      if (!listsResponse.value || listsResponse.value.length === 0) {
        throw new Error(`List "${listTitle}" not found`);
      }
      
      const listData = listsResponse.value[0];
      const listId = listData.id;
      
      // Шаг 2: Получить первый элемент списка
      const itemsResponse = await graphClient
        .api(`/sites/${this._targetSiteId}/lists/${listId}/items?$top=1&$expand=fields`)
        .get();
      
      if (!itemsResponse.value || itemsResponse.value.length === 0) {
        throw new Error(`No items found in list "${listTitle}"`);
      }
      
      const firstItem = itemsResponse.value[0];
      const itemId = firstItem.id;
      
      // Шаг 3: Обновить поле Title элемента с использованием авторизованного доступа
      this.logInfo(`Updating Title field of item #${itemId} to "${newValue}"`);
      
      await graphClient
        .api(`/sites/${this._targetSiteId}/lists/${listId}/items/${itemId}/fields`)
        .update({
          Title: newValue
        });
      
      this.logInfo(`Successfully updated Title field of item #${itemId} in list "${listTitle}" to "${newValue}"`);
      
      // Шаг 4: Получить обновленный элемент с использованием авторизованного доступа
      const updatedItemResponse = await graphClient
        .api(`/sites/${this._targetSiteId}/lists/${listId}/items/${itemId}?$expand=fields`)
        .get();
      
      // Преобразуем ответ в наш интерфейс IListItem
      const updatedItem: IListItem = {
        Id: parseInt(updatedItemResponse.id, 10),
        Title: updatedItemResponse.fields.Title,
        ...updatedItemResponse.fields
      };
      
      return updatedItem;
    } catch (error) {
      if (error.statusCode === 401 || error.statusCode === 403) {
        this.logError(`Authorization failed when updating item in list "${listTitle}". Check that Sites.ReadWrite.All permission is approved.`);
      } else {
        this.logError(`Failed to update Title field: ${error instanceof Error ? error.message : String(error)}`);
      }
      throw error;
    }
  }

  /**
   * Логирует информационное сообщение
   * @param message сообщение для логирования
   */
  protected logInfo(message: string): void {
    console.log(`[${this._logSource}] ${message}`);
  }

  /**
   * Логирует сообщение об ошибке
   * @param message сообщение об ошибке для логирования
   */
  protected logError(message: string): void {
    console.error(`[${this._logSource}] ${message}`);
  }
}