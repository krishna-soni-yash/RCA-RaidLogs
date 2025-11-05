/* eslint-disable */
import * as React from 'react';
//import PpoApproversContext from '../PpoApproversContext';
import styles from './Header.module.scss';
import { 
  Text,
  IStackTokens,
  Stack,
  Icon
} from '@fluentui/react';

export interface IHeaderProps {
  title?: string;
  subtitle?: string;
}

const Header: React.FC<IHeaderProps> = ({ 
  title = "PPO Quality Root Cause Analysis",
}) => {
//   const { approvers } = React.useContext(PpoApproversContext);
//   const contextTitle = (approvers && approvers.length > 0)
//     ? ((approvers[0] as any).Title ?? approvers[0].LinkTitle ?? title)
//     : title;
  const stackTokens: IStackTokens = { childrenGap: 6 };

  return (
  <header className={styles.header} role="banner">
      <div className={styles.inner}>
        <Stack horizontal verticalAlign="center" tokens={stackTokens}>
          <Icon iconName="Bullseye" className={styles.headerIcon} />
          <div className={styles.texts}>
            <Text className={styles.title}>{title}</Text>
          </div>
        </Stack>
      </div>
    </header>
  );
};

export default Header;
