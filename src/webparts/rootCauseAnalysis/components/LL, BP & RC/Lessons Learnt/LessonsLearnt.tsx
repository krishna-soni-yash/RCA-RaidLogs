import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';
  
interface ILessonsLearntProps {
  context: WebPartContext;
}

const LessonsLearnt: React.FC<ILessonsLearntProps> = ({ context }) => {
  return (
    <div>
        Lessons Learnt Content
    </div>
  );
}

export default LessonsLearnt;