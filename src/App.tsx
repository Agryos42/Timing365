import * as React from 'react';
import { PrimaryButton } from '@fluentui/react';

export const App: React.FC = () => {
  const onClick = () => {
    console.log('Hello from Timing365 Add-in!');
  };

  return (
    <div style={{ padding: '20px' }}>
      <h1>Timing365 Outlook Add-in</h1>
      <PrimaryButton onClick={onClick}>Click me</PrimaryButton>
    </div>
  );
};
