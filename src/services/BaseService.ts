// src/services/BaseService.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/lists/web";
import { IWebInfo } from "@pnp/sp/webs";
import { IListInfo as PnPListInfo } from "@pnp/sp/lists";

// Интерфейс для информации о сайте
export interface ISiteInfo {
  Id: string;
  Title: string;
  Url: string;
  Created: string;
  LastItemModifiedDate: string;
  [key: string]: unknown; // Индексная сигнатура для дополнительных полей
}

// Интерфейс для информации о списке
export interface IListInfo {
  Id: string;
  Title: string;
  ItemCount: number;
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
  protected _prevSiteSp: SPFI;
  protected _logSource: string;
  
  // URL предыдущего сайта
  protected _prevSiteUrl: string = "https://kpfaie.sharepoint.com/sites/KPFAData";
  protected _context: WebPartContext;

  constructor(context: WebPartContext, logSource: string) {
    this._context = context;
    // Инициализируем PnP JS с контекстом для текущего сайта
    this._sp = spfi().using(SPFx(context));
    this._logSource = logSource;
    
    // Сразу инициализируем SPFI для предыдущего сайта
    this._prevSiteSp = spfi(this._prevSiteUrl).using(SPFx(context));
  }

  /**
   * Получает URL предыдущего сайта
   * @returns URL предыдущего сайта
   */
  public getPrevSiteUrl(): string {
    return this._prevSiteUrl;
  }

  /**
   * Получает экземпляр SPFI для работы с предыдущим сайтом
   * @returns Экземпляр SPFI для предыдущего сайта
   */
  protected getPrevSiteSP(): SPFI {
    return this._prevSiteSp;
  }

  /**
   * Проверяет соединение с предыдущим сайтом
   * @returns Promise с информацией о веб-сайте
   */
  public async testPrevSiteConnection(): Promise<ISiteInfo> {
    try {
      const webInfo: IWebInfo = await this._prevSiteSp.web();
      this.logInfo(`Successfully connected to previous site: ${webInfo.Title}`);
      
      // Безопасное преобразование из IWebInfo в ISiteInfo
      const siteInfo: ISiteInfo = {
        Id: webInfo.Id,
        Title: webInfo.Title,
        Url: webInfo.Url,
        Created: webInfo.Created,
        LastItemModifiedDate: webInfo.LastItemModifiedDate
      };
      
      // Копируем другие свойства (только известные)
      // Вместо Object.keys и доступа по индексу используем явное копирование
      if (webInfo.Description) siteInfo.Description = webInfo.Description;
      if (webInfo.ServerRelativeUrl) siteInfo.ServerRelativeUrl = webInfo.ServerRelativeUrl;
      if (webInfo.WebTemplate) siteInfo.WebTemplate = webInfo.WebTemplate;
      // Можно добавить другие нужные поля
      
      return siteInfo;
    } catch (error) {
      this.logError(`Failed to connect to previous site: ${error instanceof Error ? error.message : String(error)}`);
      throw error;
    }
  }

  /**
   * Проверяет доступность списка на предыдущем сайте
   * @param listTitle Название списка для проверки
   * @returns Promise с информацией о списке или ошибкой
   */
  public async checkListExists(listTitle: string): Promise<IListInfo> {
    try {
      const pnpListInfo: PnPListInfo = await this._prevSiteSp.web.lists
        .getByTitle(listTitle)
        .select('Id,Title,ItemCount')();
      
      this.logInfo(`Successfully accessed list "${listTitle}" with ${pnpListInfo.ItemCount} items`);
      
      // Безопасное преобразование из PnP IListInfo в наш интерфейс IListInfo
      const listInfo: IListInfo = {
        Id: pnpListInfo.Id,
        Title: pnpListInfo.Title,
        ItemCount: pnpListInfo.ItemCount
      };
      
      // Копируем другие свойства (только известные)
      // Вместо Object.keys и доступа по индексу используем явное копирование
      if (pnpListInfo.Description) listInfo.Description = pnpListInfo.Description;
      if (pnpListInfo.DefaultViewUrl) listInfo.DefaultViewUrl = pnpListInfo.DefaultViewUrl;
      if (pnpListInfo.LastItemModifiedDate) listInfo.LastItemModifiedDate = pnpListInfo.LastItemModifiedDate;
      // Можно добавить другие нужные поля
      
      return listInfo;
    } catch (error) {
      this.logError(`Failed to access list "${listTitle}": ${error instanceof Error ? error.message : String(error)}`);
      throw error;
    }
  }

  /**
   * Проверяет все необходимые списки на предыдущем сайте
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
   * Обновляет значение поля Title в первом элементе списка
   * @param listTitle Название списка
   * @param newValue Новое значение для поля Title
   * @returns Promise с обновленным элементом
   */
  public async updateFirstItemTitleField(listTitle: string, newValue: string): Promise<IListItem> {
    try {
      // Получаем первый элемент списка
      const items = await this._prevSiteSp.web.lists.getByTitle(listTitle)
        .items.top(1)();
      
      if (items.length === 0) {
        throw new Error(`No items found in list "${listTitle}"`);
      }
      
      const firstItem = items[0];
      const itemId = firstItem.Id;
      
      // Обновляем поле Title элемента
      await this._prevSiteSp.web.lists.getByTitle(listTitle)
        .items.getById(itemId).update({
          Title: newValue
        });
      
      this.logInfo(`Successfully updated Title field of item #${itemId} in list "${listTitle}" to "${newValue}"`);
      
      // Получаем обновленный элемент
      const updatedItem = await this._prevSiteSp.web.lists.getByTitle(listTitle)
        .items.getById(itemId)();
      
      return updatedItem as IListItem;
    } catch (error) {
      this.logError(`Failed to update Title field: ${error instanceof Error ? error.message : String(error)}`);
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