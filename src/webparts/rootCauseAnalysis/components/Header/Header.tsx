/* eslint-disable */
import * as React from 'react';
import PpoApproversContext from '../PpoApproversContext';
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
  title = "SharePoint lists"
}) => {
  const { approver } = React.useContext(PpoApproversContext);
  const contextTitle = approver?.InternalProjectName || approver?.Title || title;
  const stackTokens: IStackTokens = { childrenGap: 6 };

  return (
  <header className={styles.header} role="banner">
      <div className={styles.inner}>
        <Stack horizontal verticalAlign="center" tokens={stackTokens} className={styles.headerContent}>
          <Icon iconName="Bullseye" className={styles.headerIcon} />
          <div className={styles.texts}>
            <Text className={styles.title}>{contextTitle}</Text>
          </div>
        </Stack>
      </div>
    </header>
  );
};

export default Header;
