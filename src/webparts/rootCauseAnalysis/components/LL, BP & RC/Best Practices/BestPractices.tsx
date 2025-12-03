import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';

interface IBestPracticesProps {
  context: WebPartContext;
}

const BestPractices: React.FC<IBestPracticesProps> = ({ context }) => {
  return (
    <div>
        Best Practices Content
    </div>
  );
}

export default BestPractices;