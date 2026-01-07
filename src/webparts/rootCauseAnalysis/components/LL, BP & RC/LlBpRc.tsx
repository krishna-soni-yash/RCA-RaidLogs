/*eslint-disable*/
import * as React from "react";
import { Pivot, PivotItem } from "@fluentui/react";
import LessonsLearnt from "./Lessons Learnt/LessonsLearnt";
import BestPractices from "./Best Practices/BestPractices";
import ReusableComponents from "./Reusable Components/ReusableComponents";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface ILlBpRcProps {
  context: WebPartContext;
}

const LlBpRc: React.FC<ILlBpRcProps> = ({ context }) => {
    return (
    <div>
      <Pivot aria-label="Lessons learnt tabs" defaultSelectedKey="lessonsLearnt">
        <PivotItem headerText="Lessons Learnt" itemKey="lessonsLearnt">
          <LessonsLearnt context={context} />
        </PivotItem>
        <PivotItem headerText="Best Practices" itemKey="bestPractices">
          <BestPractices context={context} />
        </PivotItem>
        <PivotItem headerText="Reusable Components" itemKey="reusableComponents">
          <ReusableComponents context={context} />
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default LlBpRc;