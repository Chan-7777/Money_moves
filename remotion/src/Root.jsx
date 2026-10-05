import React from 'react';
import { Composition } from 'remotion';
import { MoneyMovesFlow } from './MoneyMovesFlow';

export const RemotionRoot = () => {
  return (
    <>
      <Composition
        id="MoneyMovesFlow"
        component={MoneyMovesFlow}
        durationInFrames={750}
        fps={30}
        width={1920}
        height={1080}
      />
    </>
  );
};
