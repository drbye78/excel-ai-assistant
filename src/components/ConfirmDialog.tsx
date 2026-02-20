// Confirm Dialog Component - Standardized confirmation dialogs
// For delete confirmations, unsaved changes, etc.

import React from 'react';
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  Stack,
  Text,
  Icon,
} from '@fluentui/react';

export type ConfirmType = 'delete' | 'warning' | 'info' | 'unsaved';

export interface ConfirmDialogProps {
  /** Whether dialog is open */
  isOpen: boolean;
  /** Dialog type (determines icon/color) */
  type?: ConfirmType;
  /** Dialog title */
  title: string;
  /** Dialog message/content */
  message: string | React.ReactNode;
  /** Confirm button text */
  confirmText?: string;
  /** Cancel button text */
  cancelText?: string;
  /** Callback when confirmed */
  onConfirm: () => void;
  /** Callback when cancelled/closed */
  onCancel: () => void;
  /** Whether confirm action is destructive (delete, etc.) */
  isDestructive?: boolean;
  /** Additional details/footnote */
  details?: string;
}

/**
 * Confirm Dialog Component
 * 
 * Standardized confirmation dialogs for common scenarios.
 * Ensures consistent messaging and button placement.
 * 
 * Usage:
 * ```tsx
 * const [showDelete, setShowDelete] = useState(false);
 * 
 * <ConfirmDialog
 *   isOpen={showDelete}
 *   type="delete"
 *   title="Delete Recipe"
 *   message="Are you sure you want to delete 'Monthly Report'?"
 *   confirmText="Delete"
 *   onConfirm={handleDelete}
 *   onCancel={() => setShowDelete(false)}
 *   isDestructive
 * />
 * ```
 */
export const ConfirmDialog: React.FC<ConfirmDialogProps> = ({
  isOpen,
  type = 'info',
  title,
  message,
  confirmText,
  cancelText = 'Cancel',
  onConfirm,
  onCancel,
  isDestructive,
  details,
}) => {
  const config = {
    delete: {
      icon: 'Delete',
      iconColor: '#d13438',
      confirmText: confirmText || 'Delete',
      confirmButtonStyle: { backgroundColor: '#d13438', borderColor: '#d13438' },
    },
    warning: {
      icon: 'Warning',
      iconColor: '#ffc107',
      confirmText: confirmText || 'Continue',
      confirmButtonStyle: {},
    },
    info: {
      icon: 'Info',
      iconColor: '#0078d4',
      confirmText: confirmText || 'OK',
      confirmButtonStyle: {},
    },
    unsaved: {
      icon: 'Save',
      iconColor: '#0078d4',
      confirmText: confirmText || 'Save',
      cancelText: 'Don\'t Save',
      confirmButtonStyle: {},
    },
  };

  const currentConfig = config[type];
  const effectiveConfirmText = confirmText || currentConfig.confirmText;
  const effectiveCancelText = type === 'unsaved' ? currentConfig.cancelText : cancelText;

  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onCancel}
      dialogContentProps={{
        type: DialogType.normal,
        title,
        showCloseButton: true,
      }}
      modalProps={{
        isBlocking: true,
        styles: { main: { maxWidth: 450 } },
      }}
    >
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal tokens={{ childrenGap: 12 }}>
          <Icon
            iconName={currentConfig.icon}
            styles={{
              root: {
                fontSize: 32,
                color: currentConfig.iconColor,
                flexShrink: 0,
              },
            }}
          />
          <Stack tokens={{ childrenGap: 8 }}>
            {typeof message === 'string' ? (
              <Text>{message}</Text>
            ) : (
              message
            )}
            {details && (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {details}
              </Text>
            )}
          </Stack>
        </Stack>
      </Stack>

      <DialogFooter>
        <PrimaryButton
          onClick={onConfirm}
          text={effectiveConfirmText}
          styles={{
            root: currentConfig.confirmButtonStyle,
          }}
        />
        <DefaultButton onClick={onCancel} text={effectiveCancelText} />
      </DialogFooter>
    </Dialog>
  );
};

// Hook for managing confirm dialog state
export const useConfirmDialog = () => {
  const [state, setState] = React.useState<{
    isOpen: boolean;
    config?: Omit<ConfirmDialogProps, 'isOpen' | 'onConfirm' | 'onCancel'>;
    resolve?: (value: boolean) => void;
  }>({ isOpen: false });

  const confirm = React.useCallback(
    (config: Omit<ConfirmDialogProps, 'isOpen' | 'onConfirm' | 'onCancel'>): Promise<boolean> => {
      return new Promise((resolve) => {
        setState({ isOpen: true, config, resolve });
      });
    },
    []
  );

  const handleConfirm = React.useCallback(() => {
    state.resolve?.(true);
    setState({ isOpen: false });
  }, [state.resolve]);

  const handleCancel = React.useCallback(() => {
    state.resolve?.(false);
    setState({ isOpen: false });
  }, [state.resolve]);

  const dialog = state.config ? (
    <ConfirmDialog
      isOpen={state.isOpen}
      onConfirm={handleConfirm}
      onCancel={handleCancel}
      {...state.config}
    />
  ) : null;

  return { confirm, dialog };
};

export default ConfirmDialog;
