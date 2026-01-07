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
  const [defaultKey, setDefaultKey] = React.useState<string>('lessonsLearnt');
  const [openItemId, setOpenItemId] = React.useState<string | null>(null);

  React.useEffect(() => {
    try {
      const params = new URLSearchParams(window.location.search);
      const id = params.get('LlBpRcId') || params.get('llbprcid') || params.get('llbpRcId');
      if (id) {
        // ensure the LL/BP/RC pivot is selected when opening from query
        setDefaultKey('lessonsLearnt');
        setOpenItemId(id);
      }
    } catch (e) {
      // ignore
    }
  }, []);

  return (
    <div>
      <Pivot aria-label="Lessons learnt tabs" defaultSelectedKey={defaultKey}>
        <PivotItem headerText="Lessons Learnt" itemKey="lessonsLearnt">
          <LessonsLearnt context={context} openItemId={openItemId} />
        </PivotItem>
        <PivotItem headerText="Best Practices" itemKey="bestPractices">
          <BestPractices context={context} openItemId={openItemId} />
        </PivotItem>
        <PivotItem headerText="Reusable Components" itemKey="reusableComponents">
          <ReusableComponents context={context} openItemId={openItemId} />
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default LlBpRc;