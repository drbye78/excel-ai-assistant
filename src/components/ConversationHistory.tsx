import * as React from "react";
import { useState, useEffect } from "react";
import {
  getConversationStorage,
  ConversationStorage,
  Conversation,
  ConversationSummary,
  AutoSaveManager
} from "@/services/conversationStorage";
import { Message } from "@/types";
import {
  Stack,
  Text,
  DefaultButton,
  PrimaryButton,
  IconButton,
  SearchBox,
  Separator,
  TooltipHost,
  ContextualMenu,
  IContextualMenuItem,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  IStackTokens,
  List
} from "@fluentui/react";

interface ConversationHistoryProps {
  currentConversation?: Conversation | null;
  onSelectConversation?: (conversation: Conversation) => void;
  onCreateConversation?: () => void;
  onClose?: () => void;
}

const stackTokens: IStackTokens = {
  childrenGap: 10
};

export const ConversationHistory: React.FC<ConversationHistoryProps> = ({
  currentConversation,
  onSelectConversation,
  onCreateConversation,
  onClose
}) => {
  const [storage] = useState<ConversationStorage>(getConversationStorage());
  const [autoSaveManager] = useState<AutoSaveManager>(() => new AutoSaveManager(storage));
  const [conversations, setConversations] = useState<ConversationSummary[]>([]);
  const [filteredConversations, setFilteredConversations] = useState<ConversationSummary[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [contextMenuProps, setContextMenuProps] = useState<{
    items: IContextualMenuItem[];
    target: HTMLElement;
  } | null>(null);
  const [renameDialogOpen, setRenameDialogOpen] = useState(false);
  const [newTitle, setNewTitle] = useState("");
  const [conversationToRename, setConversationToRename] = useState<string | null>(null);
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const [conversationToDelete, setConversationToDelete] = useState<string | null>(null);
  const [stats, setStats] = useState<{ total: number; totalMessages: number } | null>(null);

  // Initialize storage and load conversations
  useEffect(() => {
    initializeStorage();
  }, []);

  // Set up auto-save for current conversation
  useEffect(() => {
    if (currentConversation) {
      autoSaveManager.startAutoSave(currentConversation);
    }

    return () => {
      autoSaveManager.stopAutoSave();
    };
  }, [currentConversation, autoSaveManager]);

  // Filter conversations when search query changes
  useEffect(() => {
    if (searchQuery.trim() === "") {
      setFilteredConversations(conversations);
    } else {
      const filtered = conversations.filter(
        conv =>
          conv.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
          conv.lastMessage.toLowerCase().includes(searchQuery.toLowerCase()) ||
          conv.workbookName.toLowerCase().includes(searchQuery.toLowerCase())
      );
      setFilteredConversations(filtered);
    }
  }, [searchQuery, conversations]);

  const initializeStorage = async () => {
    setIsLoading(true);
    setError(null);

    try {
      await storage.initialize();
      await loadConversations();
    } catch (err) {
      setError("Failed to initialize storage: " + err.message);
    } finally {
      setIsLoading(false);
    }
  };

  const loadConversations = async () => {
    try {
      const summaries = await storage.getConversationSummaries();
      setConversations(summaries);
      setFilteredConversations(summaries);

      // Load stats
      const storageStats = await storage.getStats();
      setStats({
        total: storageStats.totalConversations,
        totalMessages: storageStats.totalMessages
      });
    } catch (err) {
      setError("Failed to load conversations: " + err.message);
    }
  };

  const handleSelectConversation = async (id: string) => {
    setSelectedId(id);
    setIsLoading(true);

    try {
      const conversation = await storage.getConversation(id);
      if (conversation && onSelectConversation) {
        onSelectConversation(conversation);
      }
    } catch (err) {
      setError("Failed to load conversation: " + err.message);
    } finally {
      setIsLoading(false);
    }
  };

  const handleCreateConversation = async () => {
    try {
      // Get current workbook info from Excel
      let workbookId = "unknown";
      let workbookName = "Untitled";

      try {
        await Excel.run(async (context) => {
          const workbook = context.workbook;
          workbook.load("name");
          await context.sync();
          workbookName = workbook.name || "Untitled";
          workbookId = workbookName; // Use name as ID for simplicity
        });
      } catch (e) {
        // Excel not available
      }

      const newConversation = await storage.createConversation(
        workbookId,
        workbookName,
        "New Conversation"
      );

      if (onCreateConversation) {
        onCreateConversation();
      }

      if (onSelectConversation) {
        onSelectConversation(newConversation);
      }

      await loadConversations();
    } catch (err) {
      setError("Failed to create conversation: " + err.message);
    }
  };

  const handleContextMenu = (ev: React.MouseEvent, conv: ConversationSummary) => {
    ev.preventDefault();

    const menuItems: IContextualMenuItem[] = [
      {
        key: 'rename',
        text: 'Rename',
        iconProps: { iconName: 'Edit' },
        onClick: () => {
          setConversationToRename(conv.id);
          setNewTitle(conv.title);
          setRenameDialogOpen(true);
        }
      },
      {
        key: 'pin',
        text: conv.isPinned ? 'Unpin' : 'Pin',
        iconProps: { iconName: conv.isPinned ? 'Unpin' : 'Pin' },
        onClick: () => handleTogglePin(conv.id, !conv.isPinned)
      },
      {
        key: 'export',
        text: 'Export',
        iconProps: { iconName: 'Download' },
        subMenuProps: {
          items: [
            {
              key: 'export-json',
              text: 'Export as JSON',
              onClick: () => handleExport(conv.id, 'json')
            },
            {
              key: 'export-markdown',
              text: 'Export as Markdown',
              onClick: () => handleExport(conv.id, 'markdown')
            }
          ]
        }
      },
      {
        key: 'delete',
        text: 'Delete',
        iconProps: { iconName: 'Delete' },
        onClick: () => {
          setConversationToDelete(conv.id);
          setDeleteDialogOpen(true);
        }
      }
    ];

    setContextMenuProps({
      items: menuItems,
      target: ev.target as HTMLElement
    });
  };

  const handleTogglePin = async (id: string, isPinned: boolean) => {
    try {
      await storage.setPinned(id, isPinned);
      await loadConversations();
    } catch (err) {
      setError("Failed to update pin status: " + err.message);
    }
  };

  const handleRename = async () => {
    if (conversationToRename && newTitle.trim()) {
      try {
        await storage.updateTitle(conversationToRename, newTitle.trim());
        await loadConversations();
        setRenameDialogOpen(false);
        setConversationToRename(null);
        setNewTitle("");
      } catch (err) {
        setError("Failed to rename conversation: " + err.message);
      }
    }
  };

  const handleDelete = async () => {
    if (conversationToDelete) {
      try {
        await storage.deleteConversation(conversationToDelete);
        await loadConversations();
        setDeleteDialogOpen(false);
        setConversationToDelete(null);
      } catch (err) {
        setError("Failed to delete conversation: " + err.message);
      }
    }
  };

  const handleExport = async (id: string, format: 'json' | 'markdown') => {
    try {
      let content: string;
      let filename: string;
      let mimeType: string;

      if (format === 'json') {
        content = await storage.exportConversation(id);
        filename = `conversation-${id}.json`;
        mimeType = 'application/json';
      } else {
        content = await storage.exportConversationAsMarkdown(id);
        filename = `conversation-${id}.md`;
        mimeType = 'text/markdown';
      }

      // Create download
      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      setError("Failed to export conversation: " + err.message);
    }
  };

  const formatDate = (date: Date): string => {
    const now = new Date();
    const diff = now.getTime() - date.getTime();
    const days = Math.floor(diff / (1000 * 60 * 60 * 24));

    if (days === 0) {
      return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    } else if (days === 1) {
      return 'Yesterday';
    } else if (days < 7) {
      return date.toLocaleDateString([], { weekday: 'short' });
    } else {
      return date.toLocaleDateString([], { month: 'short', day: 'numeric' });
    }
  };

  const renderConversationItem = (conv: ConversationSummary): JSX.Element => {
    const isSelected = conv.id === selectedId ||
      (currentConversation && conv.id === currentConversation.id);

    return (
      <div
        key={conv.id}
        onClick={() => handleSelectConversation(conv.id)}
        onContextMenu={(e) => handleContextMenu(e, conv)}
        style={{
          padding: '12px',
          cursor: 'pointer',
          backgroundColor: isSelected ? '#e3f2fd' : 'transparent',
          borderLeft: isSelected ? '3px solid #2196f3' : '3px solid transparent',
          borderBottom: '1px solid #f0f0f0'
        }}
      >
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: 1 } }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              {conv.isPinned && (
                <span style={{ fontSize: '12px' }}>📌</span>
              )}
              <Text
                variant="medium"
                styles={{
                  root: {
                    fontWeight: 600,
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap'
                  }
                }}
              >
                {conv.title}
              </Text>
            </Stack>
            <Text
              variant="small"
              styles={{
                root: {
                  color: '#605e5c',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap'
                }
              }}
            >
              {conv.lastMessage || 'No messages'}
            </Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <Text variant="xSmall" styles={{ root: { color: '#a19f9d' } }}>
                {conv.workbookName}
              </Text>
              <Text variant="xSmall" styles={{ root: { color: '#a19f9d' } }}>
                •
              </Text>
              <Text variant="xSmall" styles={{ root: { color: '#a19f9d' } }}>
                {conv.messageCount} messages
              </Text>
            </Stack>
          </Stack>
          <Text variant="small" styles={{ root: { color: '#a19f9d', marginLeft: '8px' } }}>
            {formatDate(conv.lastMessageDate)}
          </Text>
        </Stack>
      </div>
    );
  };

  return (
    <Stack tokens={stackTokens} styles={{ root: { height: '100%', padding: '10px' } }}>
      {/* Header */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          💬 Conversations
        </Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="New"
            iconProps={{ iconName: 'Add' }}
            onClick={handleCreateConversation}
          />
          {onClose && (
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              onClick={onClose}
              title="Close"
            />
          )}
        </Stack>
      </Stack>

      {/* Stats */}
      {stats && (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {stats.total} conversations • {stats.totalMessages} messages
        </Text>
      )}

      <Separator />

      {/* Search */}
      <SearchBox
        placeholder="Search conversations..."
        value={searchQuery}
        onChange={(_, newValue) => setSearchQuery(newValue || '')}
        onClear={() => setSearchQuery('')}
      />

      {/* Error Message */}
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setError(null)}
          dismissButtonAriaLabel="Close"
        >
          {error}
        </MessageBar>
      )}

      {/* Loading */}
      {isLoading && (
        <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
          <Spinner size={SpinnerSize.small} />
          <Text>Loading...</Text>
        </Stack>
      )}

      {/* Conversation List */}
      <div style={{ flex: 1, overflowY: 'auto', marginTop: '10px' }}>
        {filteredConversations.length === 0 ? (
          <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { padding: '40px 20px' } }}>
            <Text variant="medium" styles={{ root: { color: '#605e5c' } }}>
              {searchQuery ? 'No conversations found' : 'No conversations yet'}
            </Text>
            {!searchQuery && (
              <DefaultButton
                text="Start a new conversation"
                onClick={handleCreateConversation}
              />
            )}
          </Stack>
        ) : (
          <>
            {filteredConversations
              .filter(conv => conv.isPinned)
              .map(renderConversationItem)}
            {filteredConversations.filter(conv => conv.isPinned).length > 0 &&
              filteredConversations.filter(conv => !conv.isPinned).length > 0 && (
                <Separator styles={{ root: { margin: '8px 0' } }} />
              )}
            {filteredConversations
              .filter(conv => !conv.isPinned)
              .map(renderConversationItem)}
          </>
        )}
      </div>

      {/* Context Menu */}
      {contextMenuProps && (
        <ContextualMenu
          items={contextMenuProps.items}
          target={contextMenuProps.target}
          onDismiss={() => setContextMenuProps(null)}
        />
      )}

      {/* Rename Dialog */}
      <Dialog
        hidden={!renameDialogOpen}
        onDismiss={() => setRenameDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Rename Conversation',
          subText: 'Enter a new name for this conversation'
        }}
      >
        <TextField
          value={newTitle}
          onChange={(_, value) => setNewTitle(value || '')}
          placeholder="Conversation name"
          autoFocus
        />
        <DialogFooter>
          <PrimaryButton text="Save" onClick={handleRename} disabled={!newTitle.trim()} />
          <DefaultButton text="Cancel" onClick={() => setRenameDialogOpen(false)} />
        </DialogFooter>
      </Dialog>

      {/* Delete Dialog */}
      <Dialog
        hidden={!deleteDialogOpen}
        onDismiss={() => setDeleteDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Conversation',
          subText: 'Are you sure you want to delete this conversation? This action cannot be undone.'
        }}
      >
        <DialogFooter>
          <PrimaryButton text="Delete" onClick={handleDelete} styles={{ root: { backgroundColor: '#d83b01' } }} />
          <DefaultButton text="Cancel" onClick={() => setDeleteDialogOpen(false)} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default ConversationHistory;
