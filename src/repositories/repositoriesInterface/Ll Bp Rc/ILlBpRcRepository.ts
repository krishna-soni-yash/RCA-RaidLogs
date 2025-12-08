import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ILessonsLearnt } from '../../../models/Ll Bp Rc/LessonsLearnt';
import { IBestPracticeAttachment, IBestPractices } from '../../../models/Ll Bp Rc/BestPractices';
import { IReusableComponentAttachment, IReusableComponents } from '../../../models/Ll Bp Rc/ReusableComponents';

export interface ILlBpRcRepository {
  fetchLessonsLearnt(useCache?: boolean, context?: WebPartContext): Promise<ILessonsLearnt[]>;
  addLessonsLearnt(item: ILessonsLearnt, context?: WebPartContext): Promise<ILessonsLearnt>;
  updateLessonsLearnt(item: ILessonsLearnt, context?: WebPartContext): Promise<ILessonsLearnt>;
  
  fetchBestPractices(useCache?: boolean, context?: WebPartContext): Promise<IBestPractices[]>;
  addBestPractices(item: IBestPractices, context?: WebPartContext): Promise<IBestPractices>;
  updateBestPractices(item: IBestPractices, context?: WebPartContext): Promise<IBestPractices>;
  getBestPracticeAttachments(itemId: number, context?: WebPartContext): Promise<IBestPracticeAttachment[]>;
  
  fetchReusableComponents(useCache?: boolean, context?: WebPartContext): Promise<IReusableComponents[]>;
  addReusableComponents(item: IReusableComponents, context?: WebPartContext): Promise<IReusableComponents>;
  updateReusableComponents(item: IReusableComponents, context?: WebPartContext): Promise<IReusableComponents>;
  getReusableComponentAttachments(itemId: number, context?: WebPartContext): Promise<IReusableComponentAttachment[]>;
}

export default ILlBpRcRepository;