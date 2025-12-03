import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';

interface IReusableComponentsProps {
  context: WebPartContext;
}
const ReusableComponents: React.FC<IReusableComponentsProps> = ({ context }) => {
  return (
    <div>
        Reusable Components Content
    </div>
  );
}

export default ReusableComponents;