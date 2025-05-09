// src/services/BaseService.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

// Интерфейс для информации о сайте
export interface ISiteInfo {
  Id: string;
  Title: string;
  Url: string;
  Created: string;
  LastItemModifiedDate: string;
  [key: string]: any; // Для любых других свойств
}

// Интерфейс для информации о списке
export interface IListInfo {
  Id: string;
  Title: string;
  ItemCount: number;
  [key: string]: any; // Для любых других свойств
}

// Интерфейс для результатов проверки списков
export interface IListCheckResult {
  [listName: string]: IListInfo | { error: string };
}

export class BaseService {
  protected _sp: SPFI;
  protected _prevSiteSp: SPFI;
  protected _logSource: string;
  
  // Замените на URL вашего предыдущего сайта
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
      const webInfo = await this._prevSiteSp.web();
      this.logInfo(`Successfully connected to previous site: ${webInfo.Title}`);
      return webInfo;
    } catch (error) {
      this.logError(`Failed to connect to previous site: ${error}`);
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
      const listInfo = await this._prevSiteSp.web.lists
        .getByTitle(listTitle)
        .select('Id,Title,ItemCount')();
      
      this.logInfo(`Successfully accessed list "${listTitle}" with ${listInfo.ItemCount} items`);
      return listInfo;
    } catch (error) {
      this.logError(`Failed to access list "${listTitle}": ${error}`);
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
        results[listTitle] = { error: error.message };
      }
    }
    
    return results;
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