import * as React from 'react';
import styles from './ModernPageCopy.module.scss';
import type { IModernPageCopyProps } from './IModernPageCopyProps';
import { PrimaryButton, Stack, TextField } from '@fluentui/react';

export const ModernPageCopy = (props: IModernPageCopyProps) => {
  // Set up Variables
  const {
    copyPage,
    fieldTitle
  } = props;
  const [pageName, setPageName] = React.useState("");



  return (
    <section className={`${styles.modernPageCopy}`}>
      <Stack enableScopedSelectors tokens={{ childrenGap: 5 }}>
        <TextField label={fieldTitle} value={pageName} onChange={((event: React.FormEvent<HTMLInputElement>, text: string) => { setPageName(text) })} />
        <PrimaryButton text="Create" onClick={() => copyPage(pageName)} allowDisabledFocus />
      </Stack>
    </section>
  );
}

export default ModernPageCopy;