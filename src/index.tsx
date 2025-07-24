import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import { App } from './App';

initializeIcons();

ReactDOM.render(<App />, document.getElementById('root'));
