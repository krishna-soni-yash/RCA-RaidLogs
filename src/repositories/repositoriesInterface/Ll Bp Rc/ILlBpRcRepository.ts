import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ILessonsLearnt } from '../../../models/Ll Bp Rc/LessonsLearnt';
import { IBestPractices } from '../../../models/Ll Bp Rc/BestPractices';
import { IReusableComponents } from '../../../models/Ll Bp Rc/ReusableComponents';

export interface ILlBpRcRepository {
  fetchLessonsLearnt(useCache?: boolean, context?: WebPartContext): Promise<ILessonsLearnt[]>;
  fetchBestPractices(useCache?: boolean, context?: WebPartContext): Promise<IBestPractices[]>;
  fetchReusableComponents(useCache?: boolean, context?: WebPartContext): Promise<IReusableComponents[]>;
}

export default ILlBpRcRepository;