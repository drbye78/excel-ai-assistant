import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  Message,
  AIAction,
  AISettings,
  ExcelContext
} from "@/types";
import AIService from "@/services/aiService";
import ActionHandler from "@/services/actionHandler";
import ExcelService from "@/services/excelService";
import { useTranslation } from "./TranslationProvider";
import {
  Stack,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Dialog,
  DialogType,
  DialogFooter,
  Text,
  Separator,
  IStackTokens
} from "@fluentui/react";

interface ChatProps {
  settings: AISettings;
}

const stackTokens: IStackTokens = {
  childrenGap: 10
};

export const Chat: React.FC<ChatProps> = ({ settings }) => {
  const { t } = useTranslation();
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [pendingActions, setPendingActions] = useState<AIAction[] | null>(null);
  const [suggestedPrompts, setSuggestedPrompts] = useState<string[]>([]);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  // Add welcome message on mount
  useEffect(() => {
    const welcomeMessage: Message = {
      id: "welcome",
      role: "assistant",
      content: t('chat.welcome'),
      timestamp: new Date()
    };
    setMessages([welcomeMessage]);
  }, [t]);

  const sendMessage = async () => {
    if (!input.trim() || isLoading) return;

    if (!settings.apiKey) {
      setError(t('chat.noApiKey'));
      return;
    }

    const userMessage: Message = {
      id: Date.now().toString(),
      role: "user",
      content: input,
      timestamp: new Date()
    };

    setMessages((prev) => [...prev, userMessage]);
    setInput("");
    setIsLoading(true);
    setError(null);

    try {
      // Get current Excel context
      const excelContext = await ExcelService.getFullContext();

      // Send to AI
      const response = await AIService.sendMessage({
        message: userMessage.content,
        context: excelContext,
        conversationHistory: messages.filter((m) => m.id !== "welcome"),
        settings
      });

      const assistantMessage: Message = {
        id: (Date.now() + 1).toString(),
        role: "assistant",
        content: response.message,
        timestamp: new Date(),
        actions: response.actions
      };

      setMessages((prev) => [...prev, assistantMessage]);

      if (response.suggestedPrompts) {
        setSuggestedPrompts(response.suggestedPrompts);
      }

      // Check if actions require confirmation
      if (response.requiresConfirmation && response.actions) {
        setPendingActions(response.actions);
      } else if (response.actions && response.actions.length > 0) {
        // Auto-execute non-destructive actions
        await executeActions(response.actions);
      }
    } catch (err: unknown) {
      const errorMessage = err instanceof Error ? err.message : String(err);
      setError(errorMessage || t('error.generic'));
    } finally {
      setIsLoading(false);
    }
  };

  const executeActions = async (actions: AIAction[]) => {
    const actionResults: string[] = [];

    for (const action of actions) {
      try {
        const result = await ActionHandler.executeAction(action);
        actionResults.push(result);
      } catch (err: unknown) {
        const errMsg = err instanceof Error ? err.message : String(err);
        actionResults.push(`Error: ${errMsg}`);
      }
    }

    // Add results as system message
    const resultMessage: Message = {
      id: Date.now().toString(),
      role: "system",
      content: t('chat.actionsCompleted') + "\n" + actionResults.map((r) => `✓ ${r}`).join("\n"),
      timestamp: new Date()
    };

    setMessages((prev) => [...prev, resultMessage]);
  };

  const confirmActions = async () => {
    if (pendingActions) {
      await executeActions(pendingActions);
      setPendingActions(null);
    }
  };

  const cancelActions = () => {
    setPendingActions(null);
  };

  const handleKeyPress = (ev: React.KeyboardEvent) => {
    if (ev.key === "Enter" && !ev.shiftKey) {
      ev.preventDefault();
      sendMessage();
    }
  };

  const handleSuggestedPrompt = (prompt: string) => {
    setInput(prompt);
  };

  const handleActionClick = async (action: AIAction) => {
    try {
      const result = await ActionHandler.executeAction(action);

      const resultMessage: Message = {
        id: Date.now().toString(),
        role: "system",
        content: `✓ ${result}`,
        timestamp: new Date()
      };

      setMessages((prev) => [...prev, resultMessage]);
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : String(err);
      setError(t('chat.actionFailed') + ` ${errMsg}`);
    }
  };

  return (
    <Stack tokens={stackTokens} styles={{ root: { height: "100%", padding: "10px" } }}>
      {/* Error Message */}
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setError(null)}
          dismissButtonAriaLabel={t('common.close')}
        >
          {error}
        </MessageBar>
      )}

      {/* Messages Area */}
      <div
        style={{
          flex: 1,
          overflowY: "auto",
          border: "1px solid #e0e0e0",
          borderRadius: "4px",
          padding: "10px",
          backgroundColor: "#fafafa"
        }}
      >
        {messages.map((message) => (
          <div
            key={message.id}
            style={{
              marginBottom: "15px",
              padding: "10px",
              borderRadius: "8px",
              backgroundColor:
                message.role === "user"
                  ? "#e3f2fd"
                  : message.role === "system"
                  ? "#f3e5f5"
                  : "#ffffff",
              border:
                message.role === "user"
                  ? "1px solid #bbdefb"
                  : message.role === "system"
                  ? "1px solid #e1bee7"
                  : "1px solid #e0e0e0",
              marginLeft: message.role === "user" ? "20px" : "0",
              marginRight: message.role === "assistant" ? "20px" : "0"
            }}
          >
            <Text
              variant="small"
              styles={{
                root: {
                  fontWeight: 600,
                  color:
                    message.role === "user"
                      ? "#1565c0"
                      : message.role === "system"
                      ? "#7b1fa2"
                      : "#2e7d32",
                  marginBottom: "5px",
                  display: "block"
                }
              }}
            >
              {message.role === "user" ? t('chat.role.user') : message.role === "system" ? t('chat.role.system') : t('chat.role.assistant')}
            </Text>
            <Text
              styles={{
                root: {
                  whiteSpace: "pre-wrap",
                  wordBreak: "break-word"
                }
              }}
            >
              {message.content}
            </Text>

            {/* Action Buttons */}
            {message.actions && message.actions.length > 0 && (
              <Stack horizontal tokens={{ childrenGap: 5 }} style={{ marginTop: "10px" }}>
                {message.actions.map((action, idx) => (
                  <DefaultButton
                    key={idx}
                    text={action.label}
                    onClick={() => handleActionClick(action)}
                    styles={{
                      root: {
                        fontSize: "12px",
                        height: "24px"
                      }
                    }}
                  />
                ))}
              </Stack>
            )}
          </div>
        ))}
        <div ref={messagesEndRef} />
      </div>

      {/* Suggested Prompts */}
      {suggestedPrompts.length > 0 && (
        <div>
          <Text variant="small" styles={{ root: { color: "#666", marginBottom: "5px" } }}>
            {t('chat.suggested')}
          </Text>
          <Stack horizontal wrap tokens={{ childrenGap: 5 }}>
            {suggestedPrompts.map((prompt, idx) => (
              <DefaultButton
                key={idx}
                text={prompt}
                onClick={() => handleSuggestedPrompt(prompt)}
                styles={{
                  root: {
                    fontSize: "11px",
                    height: "22px"
                  }
                }}
              />
            ))}
          </Stack>
        </div>
      )}

      <Separator />

      {/* Input Area */}
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <TextField
          value={input}
          onChange={(_, newValue) => setInput(newValue || "")}
          onKeyPress={handleKeyPress}
          placeholder={t('chat.placeholder')}
          multiline
          rows={2}
          resizable={false}
          styles={{
            root: { flex: 1 },
            fieldGroup: {
              minHeight: "60px"
            }
          }}
          disabled={isLoading}
        />
        <PrimaryButton
          text={isLoading ? "..." : t('chat.send')}
          onClick={sendMessage}
          disabled={isLoading || !input.trim()}
          styles={{
            root: {
              alignSelf: "flex-end",
              height: "60px",
              minWidth: "80px"
            }
          }}
        >
          {isLoading && <Spinner size={SpinnerSize.small} />}
        </PrimaryButton>
      </Stack>

      {/* Confirmation Dialog */}
      <Dialog
        hidden={!pendingActions}
        onDismiss={cancelActions}
        dialogContentProps={{
          type: DialogType.normal,
          title: t('chat.confirm.title'),
          subText: t('chat.confirm.message')
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={confirmActions} text={t('common.ok')} />
          <DefaultButton onClick={cancelActions} text={t('common.cancel')} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};
