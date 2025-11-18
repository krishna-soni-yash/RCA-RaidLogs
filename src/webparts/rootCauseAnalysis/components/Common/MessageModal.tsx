import * as React from 'react';
import { Modal, IconButton, PrimaryButton, Icon } from '@fluentui/react';
import styles from './MessageModal.module.scss';

export type MessageType = 'success' | 'error' | 'warning' | 'info';

export interface IMessageModalProps {
  isOpen: boolean;
  message: string;
  type: MessageType;
  onDismiss: () => void;
}

const MessageModal: React.FC<IMessageModalProps> = ({ isOpen, message, type, onDismiss }) => {
  const getIcon = (): string => {
    switch (type) {
      case 'success':
        return 'CheckMark';
      case 'error':
        return 'ErrorBadge';
      case 'warning':
        return 'Warning';
      case 'info':
      default:
        return 'Info';
    }
  };

  const getIconColor = (): string => {
    switch (type) {
      case 'success':
        return '#107C10';
      case 'error':
        return '#A80000';
      case 'warning':
        return '#F7630C';
      case 'info':
      default:
        return '#0078D4';
    }
  };

  const getTitle = (): string => {
    switch (type) {
      case 'success':
        return 'Success';
      case 'error':
        return 'Error';
      case 'warning':
        return 'Warning';
      case 'info':
      default:
        return 'Information';
    }
  };

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={styles.messageModalContainer}
    >
      <div className={styles.messageModalContent}>
        <div className={styles.messageModalHeader}>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Close"
            onClick={onDismiss}
            className={styles.closeButton}
          />
        </div>
        <div className={styles.messageModalBody}>
          <div className={styles.iconContainer}>
            <Icon
              iconName={getIcon()}
              className={styles.icon}
              style={{ color: getIconColor(), fontSize: '48px' }}
            />
          </div>
          <h2 className={styles.title}>{getTitle()}</h2>
          <p className={styles.message}>{message}</p>
        </div>
        <div className={styles.messageModalFooter}>
          <PrimaryButton text="OK" onClick={onDismiss} />
        </div>
      </div>
    </Modal>
  );
};

export default MessageModal;
