/* eslint-disable */
import * as React from 'react';
//import PpoApproversContext from '../PpoApproversContext';
import styles from './Header.module.scss';
import { 
  Text,
  IStackTokens,
  Stack,
  Icon,
  Pivot,
  PivotItem
} from '@fluentui/react';

export interface IHeaderProps {
  title?: string;
  subtitle?: string;
  activeTab?: string;
  onTabChange?: (tabKey: string) => void;
}

const Header: React.FC<IHeaderProps> = ({ 
  title = "SharePoint lists",
  activeTab = "rootCauseAnalysis",
  onTabChange
}) => {
//   const { approvers } = React.useContext(PpoApproversContext);
//   const contextTitle = (approvers && approvers.length > 0)
//     ? ((approvers[0] as any).Title ?? approvers[0].LinkTitle ?? title)
//     : title;
  const stackTokens: IStackTokens = { childrenGap: 6 };

  const handleTabChange = (item?: PivotItem): void => {
    if (item && onTabChange) {
      onTabChange(item.props.itemKey || "rootCauseAnalysis");
    }
  };

  return (
  <header className={styles.header} role="banner">
      <div className={styles.inner}>
        <Stack horizontal verticalAlign="center" tokens={stackTokens} className={styles.headerContent}>
          <Icon iconName="Bullseye" className={styles.headerIcon} />
          <div className={styles.texts}>
            <Text className={styles.title}>{title}</Text>
          </div>
        </Stack>
        <div className={styles.tabsContainer}>
          <Pivot
            selectedKey={activeTab}
            onLinkClick={handleTabChange}
            className={styles.tabs}
          >
            <PivotItem 
              headerText="Root Cause Analysis" 
              itemKey="rootCauseAnalysis"
            />
            <PivotItem 
              headerText="Raid Logs" 
              itemKey="raidLogs"
            />
          </Pivot>
        </div>
      </div>
    </header>
  );
};

export default Header;
