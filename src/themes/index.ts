import { defaultTheme } from './default';
import { elegantTheme } from './elegant';
import { techTheme } from './tech';
import { pinkTheme } from './pink';
import type { Theme } from './types';

export const themes: Theme[] = [defaultTheme, elegantTheme, techTheme, pinkTheme];

export function getTheme(id: string): Theme {
  return themes.find((t) => t.id === id) ?? defaultTheme;
}

export type { Theme, StyleMap } from './types';
