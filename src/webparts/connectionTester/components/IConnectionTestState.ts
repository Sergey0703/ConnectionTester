import { MessageBarType } from 'office-ui-fabric-react';
import { ISiteInfo, IListCheckResult, IListItem } from '../../../services/BaseService';

export interface IConnectionTestState {
  loading: boolean;
  error: string | null;
  prevSiteUrl: string;
  siteInfo: ISiteInfo | null;
  listsCheckResult: IListCheckResult | null;
  updateStatus: {
    type: MessageBarType;
    message: string;
  } | null;
  updateTitle: string;
  listTitle: string;
  updatedItem: IListItem | null;
}