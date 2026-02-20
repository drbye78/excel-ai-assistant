// Onboarding Tour Component - Guided user introduction
// Uses TeachingBubble for step-by-step feature walkthrough

import React, { useState, useEffect, useCallback } from 'react';
import {
  TeachingBubble,
  DirectionalHint,
  PrimaryButton,
  DefaultButton,
  Stack,
  Text,
  ProgressIndicator,
} from '@fluentui/react';

export interface TourStep {
  /** Target element selector (CSS selector) */
  target: string;
  /** Step title */
  title: string;
  /** Step content/description */
  content: string;
  /** Primary button text */
  primaryButtonText?: string;
  /** Secondary button text */
  secondaryButtonText?: string;
  /** Which side to show the bubble */
  directionalHint?: DirectionalHint;
  /** Whether this step is optional (can be skipped) */
  isOptional?: boolean;
  /** Custom action when primary button clicked */
  onPrimaryAction?: () => void;
}

export interface OnboardingTourProps {
  /** Tour identifier */
  tourId: string;
  /** Array of tour steps */
  steps: TourStep[];
  /** Whether tour is running */
  isOpen: boolean;
  /** Called when tour is completed */
  onComplete: () => void;
  /** Called when tour is dismissed */
  onDismiss: () => void;
  /** Don't show again preference key */
  preferenceKey?: string;
}

const STORAGE_KEY_PREFIX = 'tour-completed-';

/**
 * Onboarding Tour Component
 * 
 * Guided walkthrough using TeachingBubble for feature discovery.
 * 
 * Usage:
 * ```tsx
 * const steps: TourStep[] = [
 *   { target: '.chat-input', title: 'Ask Questions', content: 'Type here...' },
 *   { target: '.recipes-nav', title: 'Save Recipes', content: 'Store operations...' },
 * ];
 * 
 * <OnboardingTour
 *   tourId="first-time"
 *   steps={steps}
 *   isOpen={showTour}
 *   onComplete={() => setShowTour(false)}
 *   onDismiss={() => setShowTour(false)}
 * />
 * ```
 */
export const OnboardingTour: React.FC<OnboardingTourProps> = ({
  tourId,
  steps,
  isOpen,
  onComplete,
  onDismiss,
  preferenceKey,
}) => {
  const [currentStep, setCurrentStep] = useState(0);
  const [targetElement, setTargetElement] = useState<HTMLElement | null>(null);

  const step = steps[currentStep];
  const isFirstStep = currentStep === 0;
  const isLastStep = currentStep === steps.length - 1;
  const progress = ((currentStep + 1) / steps.length) * 100;

  // Find target element when step changes
  useEffect(() => {
    if (!isOpen || !step) return;

    const findTarget = () => {
      const element = document.querySelector(step.target) as HTMLElement;
      if (element) {
        setTargetElement(element);
        // Scroll into view if needed
        element.scrollIntoView({ behavior: 'smooth', block: 'center' });
      } else {
        // Retry after a short delay if element not found
        setTimeout(findTarget, 100);
      }
    };

    findTarget();
  }, [isOpen, step, currentStep]);

  const handleNext = useCallback(() => {
    if (isLastStep) {
      handleComplete();
    } else {
      setCurrentStep((prev) => prev + 1);
    }
  }, [isLastStep]);

  const handleBack = useCallback(() => {
    if (!isFirstStep) {
      setCurrentStep((prev) => prev - 1);
    }
  }, [isFirstStep]);

  const handleComplete = useCallback(() => {
    if (preferenceKey) {
      localStorage.setItem(STORAGE_KEY_PREFIX + preferenceKey, 'true');
    }
    onComplete();
    setCurrentStep(0);
  }, [onComplete, preferenceKey]);

  const handleDismiss = useCallback(() => {
    onDismiss();
    setCurrentStep(0);
  }, [onDismiss]);

  if (!isOpen || !step || !targetElement) return null;

  return (
    <TeachingBubble
      targetElement={targetElement}
      headline={step.title}
      hasCloseButton
      onDismiss={handleDismiss}
      directionalHint={step.directionalHint || DirectionalHint.rightCenter}
      primaryButtonProps={{
        children: step.primaryButtonText || (isLastStep ? 'Finish' : 'Next'),
        onClick: () => {
          step.onPrimaryAction?.();
          handleNext();
        },
      }}
      secondaryButtonProps={
        !isFirstStep
          ? {
              children: step.secondaryButtonText || 'Back',
              onClick: handleBack,
            }
          : undefined
      }
      styles={{
        root: { maxWidth: 320 },
        body: { padding: '16px 20px' },
        bodyContent: { padding: 0 },
        header: { marginBottom: 12 },
        headline: { fontSize: 16, fontWeight: 600 },
        subText: { fontSize: 14, marginBottom: 16 },
        footer: {
          display: 'flex',
          flexDirection: 'column',
          gap: 12,
          marginTop: 16,
        },
      }}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        <Text>{step.content}</Text>
        
        {/* Progress indicator */}
        <ProgressIndicator
          percentComplete={progress / 100}
          styles={{
            root: { marginTop: 8 },
            itemName: { display: 'none' },
            itemDescription: { display: 'none' },
          }}
        />
        
        {/* Step counter */}
        <Text
          variant="small"
          styles={{ root: { color: '#605e5c', textAlign: 'center' } }}
        >
          Step {currentStep + 1} of {steps.length}
        </Text>

        {/* Skip option for optional steps */}
        {step.isOptional && (
          <DefaultButton
            text="Skip this step"
            onClick={handleNext}
            styles={{ root: { alignSelf: 'center' } }}
          />
        )}
      </Stack>
    </TeachingBubble>
  );
};

// Hook for managing tour state
export const useOnboarding = (tourId: string, preferenceKey?: string) => {
  const [isOpen, setIsOpen] = useState(false);

  const startTour = useCallback(() => {
    const hasCompleted = preferenceKey
      ? localStorage.getItem(STORAGE_KEY_PREFIX + preferenceKey) === 'true'
      : false;
    
    if (!hasCompleted) {
      setIsOpen(true);
    }
  }, [preferenceKey]);

  const restartTour = useCallback(() => {
    setIsOpen(true);
  }, []);

  const markAsCompleted = useCallback(() => {
    if (preferenceKey) {
      localStorage.setItem(STORAGE_KEY_PREFIX + preferenceKey, 'true');
    }
  }, [preferenceKey]);

  const resetTour = useCallback(() => {
    if (preferenceKey) {
      localStorage.removeItem(STORAGE_KEY_PREFIX + preferenceKey);
    }
  }, [preferenceKey]);

  return {
    isOpen,
    setIsOpen,
    startTour,
    restartTour,
    markAsCompleted,
    resetTour,
  };
};

// Pre-configured tours
export const defaultTours = {
  firstTime: [
    {
      target: '.chat-input',
      title: 'Ask Excel Questions',
      content: 'Type natural language commands like "Create a pivot table from Sales data"',
      directionalHint: DirectionalHint.topCenter,
    },
    {
      target: '.recipes-nav',
      title: 'Save Your Recipes',
      content: 'Save common operations as reusable recipes for quick access.',
      directionalHint: DirectionalHint.rightCenter,
    },
    {
      target: '.analytics-nav',
      title: 'Track Your Usage',
      content: 'View analytics about your Excel productivity and feature usage.',
      directionalHint: DirectionalHint.rightCenter,
    },
    {
      target: '.help-button',
      title: 'Get Help Anytime',
      content: 'Click the help icon for contextual documentation and tips.',
      directionalHint: DirectionalHint.bottomCenter,
    },
  ] as TourStep[],

  recipesFeature: [
    {
      target: '.new-recipe-button',
      title: 'Create a Recipe',
      content: 'Start building a reusable operation sequence.',
    },
    {
      target: '.recipe-builder',
      title: 'Configure Steps',
      content: 'Add and arrange the operations in your recipe.',
    },
    {
      target: '.recipe-save',
      title: 'Save and Share',
      content: 'Save your recipe and optionally share with your team.',
    },
  ] as TourStep[],
};

export default OnboardingTour;
