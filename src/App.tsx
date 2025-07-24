import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Dropdown,
  IDropdownOption,
  TextField,
  PrimaryButton,
  Stack,
} from '@fluentui/react';

declare const Office: any; // Office.js global

const blocOptions: IDropdownOption[] = [
  { key: 'bloc1', text: 'Bloc 1' },
  { key: 'bloc2', text: 'Bloc 2' },
];

const subBlocOptions: IDropdownOption[] = [
  { key: 'sub1', text: 'Sous-bloc 1' },
  { key: 'sub2', text: 'Sous-bloc 2' },
];

export const App: React.FC = () => {
  const [bloc, setBloc] = useState<string | undefined>();
  const [subBloc, setSubBloc] = useState<string | undefined>();
  const [comment, setComment] = useState('');
  const [isRunning, setIsRunning] = useState(false);
  const [startTime, setStartTime] = useState<Date | null>(null);
  const [elapsed, setElapsed] = useState(0);

  useEffect(() => {
    let timer: NodeJS.Timeout | undefined;
    if (isRunning && startTime) {
      timer = setInterval(() => {
        setElapsed(Math.floor((Date.now() - startTime.getTime()) / 1000));
      }, 1000);
    }
    return () => {
      if (timer) {
        clearInterval(timer);
      }
    };
  }, [isRunning, startTime]);

  const createEvent = (start: Date) => {
    try {
      if (Office?.context?.mailbox?.item?.calendarItem) {
        // Placeholder Outlook event creation
        console.log('Creating Outlook event');
      }
    } catch (err) {
      console.error('Failed to create event', err);
    }
  };

  const startTimer = () => {
    if (isRunning) return;
    const start = new Date();
    setStartTime(start);
    setElapsed(0);
    setIsRunning(true);
    createEvent(start);
  };

  const stopTimer = () => {
    if (!isRunning) return;
    const end = new Date();
    setIsRunning(false);
    console.log('Event stopped', { bloc, subBloc, comment, start: startTime, end });
  };

  return (
    <div style={{ padding: '20px', width: 300 }}>
      <h1>Timing365</h1>
      <Stack tokens={{ childrenGap: 10 }}>
        <Dropdown
          label="Bloc"
          options={blocOptions}
          selectedKey={bloc}
          onChange={(_, o) => setBloc(o?.key as string)}
        />
        <Dropdown
          label="Sous-bloc"
          options={subBlocOptions}
          selectedKey={subBloc}
          onChange={(_, o) => setSubBloc(o?.key as string)}
        />
        <TextField
          label="Commentaire"
          multiline
          value={comment}
          onChange={(_, v) => setComment(v || '')}
        />
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton text="Démarrer" onClick={startTimer} disabled={isRunning} />
          <PrimaryButton text="Arrêter" onClick={stopTimer} disabled={!isRunning} />
        </Stack>
        <div style={{ marginTop: 10 }}>
          {isRunning
            ? `Activité en cours depuis ${elapsed}s`
            : 'Aucune activité en cours'}
        </div>
      </Stack>
    </div>
  );
};
