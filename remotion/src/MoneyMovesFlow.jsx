import React from 'react';
import {
  AbsoluteFill,
  interpolate,
  spring,
  staticFile,
  useCurrentFrame,
  useVideoConfig,
  Img,
  Sequence,
} from 'remotion';

const COLORS = {
  navy: '#0A192F',
  accent: '#005A9C',
  orange: '#E65100',
  red: '#D64545',
  blue: '#0284C7',
  cream: '#F8F9FA',
  white: '#FFFFFF',
  ink: '#112233',
  inkSoft: '#556677',
  border: '#E2E8F0',
};

const StageOverlay = ({ stageNum, title, rule, tag, tagColor, frame, fps, delay = 5 }) => {
  const enterSpring = spring({
    frame: frame - delay,
    fps,
    config: { damping: 14, mass: 0.8 },
  });

  const opacity = interpolate(enterSpring, [0, 1], [0, 1]);
  const translateX = interpolate(enterSpring, [0, 1], [-30, 0]);

  return (
    <div
      style={{
        position: 'absolute',
        top: 90,
        left: 70,
        width: 440,
        backgroundColor: 'rgba(255, 255, 255, 0.96)',
        backdropFilter: 'blur(12px)',
        border: `2px solid ${COLORS.navy}`,
        borderRadius: 18,
        padding: '24px 28px',
        boxShadow: '0 20px 40px rgba(10, 25, 47, 0.08)',
        opacity,
        transform: `translateX(${translateX}px)`,
        fontFamily: 'system-ui, -apple-system, sans-serif',
        zIndex: 50,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10 }}>
        <span
          style={{
            backgroundColor: tagColor || COLORS.navy,
            color: COLORS.white,
            fontWeight: 800,
            fontSize: 12,
            letterSpacing: '0.08em',
            textTransform: 'uppercase',
            padding: '4px 10px',
            borderRadius: 999,
          }}
        >
          {tag || `Step ${stageNum}`}
        </span>
        <span style={{ fontSize: 13, fontWeight: 700, color: COLORS.inkSoft }}>
          MoneyMoves AU Roadmap
        </span>
      </div>
      <h2
        style={{
          margin: 0,
          fontSize: 26,
          fontWeight: 800,
          color: COLORS.navy,
          letterSpacing: '-0.02em',
          lineHeight: 1.2,
        }}
      >
        {title}
      </h2>
      <p
        style={{
          margin: '12px 0 0 0',
          fontSize: 16,
          color: COLORS.inkSoft,
          lineHeight: 1.45,
          fontWeight: 500,
        }}
      >
        {rule}
      </p>
    </div>
  );
};

const VideoHud = ({ frame, totalFrames }) => {
  const progress = Math.min(100, Math.max(0, (frame / totalFrames) * 100));

  return (
    <>
      <div
        style={{
          position: 'absolute',
          top: 0,
          left: 0,
          right: 0,
          height: 52,
          padding: '0 70px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          borderBottom: `1px solid ${COLORS.border}`,
          backgroundColor: 'rgba(255, 255, 255, 0.92)',
          backdropFilter: 'blur(8px)',
          fontFamily: 'system-ui, -apple-system, sans-serif',
          zIndex: 100,
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div
            style={{
              width: 14,
              height: 14,
              borderRadius: '50%',
              backgroundColor: COLORS.navy,
            }}
          />
          <span style={{ fontWeight: 800, fontSize: 18, color: COLORS.navy, letterSpacing: '-0.02em' }}>
            MoneyMoves <span style={{ color: COLORS.accent }}>AU</span>
          </span>
        </div>
        <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.inkSoft, letterSpacing: '0.05em' }}>
          AUSTRALIAN PERSONAL FINANCE DECISION ENGINE
        </div>
      </div>

      <div
        style={{
          position: 'absolute',
          bottom: 0,
          left: 0,
          right: 0,
          height: 6,
          backgroundColor: '#E2E8F0',
          zIndex: 100,
        }}
      >
        <div
          style={{
            height: '100%',
            width: `${progress}%`,
            backgroundColor: COLORS.accent,
          }}
        />
      </div>
    </>
  );
};

export const MoneyMovesFlow = () => {
  const frame = useCurrentFrame();
  const { fps, durationInFrames } = useVideoConfig();

  return (
    <AbsoluteFill style={{ backgroundColor: COLORS.white }}>
      <VideoHud frame={frame} totalFrames={durationInFrames} />

      {/* ── Scene 0: Overview Panorama (Frames 0 - 90 / 0s - 3s) ── */}
      <Sequence from={0} durationInFrames={90}>
        <ScenePanorama frame={frame} fps={fps} />
      </Sequence>

      {/* ── Scene 1: Step 1 Inflow Splitter (Frames 90 - 240 / 3s - 8s) ── */}
      <Sequence from={90} durationInFrames={150}>
        <SceneStepOne frame={frame - 90} fps={fps} />
      </Sequence>

      {/* ── Scene 2: Step 2 Emergency Buffer (Frames 240 - 390 / 8s - 13s) ── */}
      <Sequence from={240} durationInFrames={150}>
        <SceneStepTwo frame={frame - 240} fps={fps} />
      </Sequence>

      {/* ── Scene 3: Step 3 Debt Avalanche (Frames 390 - 540 / 13s - 18s) ── */}
      <Sequence from={390} durationInFrames={150}>
        <SceneStepThree frame={frame - 390} fps={fps} />
      </Sequence>

      {/* ── Scene 4: Step 4 Stress-Testing & Growth (Frames 540 - 690 / 18s - 23s) ── */}
      <Sequence from={540} durationInFrames={150}>
        <SceneStepFour frame={frame - 540} fps={fps} />
      </Sequence>

      {/* ── Scene 5: Outro Summary Dossier (Frames 690 - 750 / 23s - 25s) ── */}
      <Sequence from={690} durationInFrames={60}>
        <SceneOutro frame={frame - 690} fps={fps} />
      </Sequence>
    </AbsoluteFill>
  );
};

const ScenePanorama = ({ frame, fps }) => {
  const zoom = interpolate(frame, [0, 90], [1.08, 1.0], { extrapolateRight: 'clamp' });
  const titleSpring = spring({ frame, fps, config: { damping: 15 } });
  const titleOpacity = interpolate(titleSpring, [0, 1], [0, 1]);
  const titleY = interpolate(titleSpring, [0, 1], [30, 0]);

  return (
    <AbsoluteFill style={{ overflow: 'hidden', justifyContent: 'center', alignItems: 'center' }}>
      <Img
        src={staticFile('illustrations/00-full-pipeline-panorama.png')}
        style={{
          width: '90%',
          height: 'auto',
          transform: `scale(${zoom}) translateY(-10px)`,
          transformOrigin: 'center center',
          filter: 'drop-shadow(0 10px 30px rgba(0,0,0,0.04))',
        }}
      />
      <div
        style={{
          position: 'absolute',
          bottom: 50,
          backgroundColor: 'rgba(10, 25, 47, 0.95)',
          padding: '18px 44px',
          borderRadius: 18,
          textAlign: 'center',
          boxShadow: '0 20px 50px rgba(0,0,0,0.2)',
          opacity: titleOpacity,
          transform: `translateY(${titleY}px)`,
          fontFamily: 'system-ui, -apple-system, sans-serif',
        }}
      >
        <h1 style={{ margin: 0, fontSize: 32, color: COLORS.white, fontWeight: 800 }}>
          The Australian Personal Finance Machine
        </h1>
        <p style={{ margin: '6px 0 0 0', fontSize: 17, color: '#94A3B8' }}>
          Stop guessing where money goes. Follow the exact order of operations.
        </p>
      </div>
    </AbsoluteFill>
  );
};

const SceneStepOne = ({ frame, fps }) => {
  const scale = interpolate(frame, [0, 150], [1.02, 1.0], { extrapolateRight: 'clamp' });

  return (
    <AbsoluteFill style={{ display: 'flex', flexDirection: 'row', alignItems: 'center' }}>
      <StageOverlay
        stageNum="1"
        tag="Stage 1 • Triage"
        tagColor={COLORS.orange}
        title="The Cash Flow Splitter"
        rule="Route every dollar immediately into Essentials (50%), Debt (20%), and Buffer (10%) while clamping discretionary leaks."
        frame={frame}
        fps={fps}
      />
      <div style={{ flex: 1, display: 'flex', justifyContent: 'flex-end', paddingRight: 60, marginTop: 40 }}>
        <Img
          src={staticFile('illustrations/01-cash-flow-splitter.png')}
          style={{
            width: '72%',
            height: 'auto',
            transform: `scale(${scale})`,
            transformOrigin: 'center center',
          }}
        />
      </div>
    </AbsoluteFill>
  );
};

const SceneStepTwo = ({ frame, fps }) => {
  const scale = interpolate(frame, [0, 150], [1.02, 1.0], { extrapolateRight: 'clamp' });

  return (
    <AbsoluteFill style={{ display: 'flex', flexDirection: 'row', alignItems: 'center' }}>
      <StageOverlay
        stageNum="2"
        tag="Stage 2 • Security"
        tagColor={COLORS.blue}
        title="Emergency Buffer Reservoir"
        rule="Lock in a 3-month survival fund mounted on shock absorbers. Protects you so unexpected life events never force panic borrowing."
        frame={frame}
        fps={fps}
      />
      <div style={{ flex: 1, display: 'flex', justifyContent: 'flex-end', paddingRight: 60, marginTop: 40 }}>
        <Img
          src={staticFile('illustrations/02-emergency-buffer-reservoir.png')}
          style={{
            width: '72%',
            height: 'auto',
            transform: `scale(${scale})`,
            transformOrigin: 'center center',
          }}
        />
      </div>
    </AbsoluteFill>
  );
};

const SceneStepThree = ({ frame, fps }) => {
  const shake =
    frame > 25 && frame < 45
      ? Math.sin(frame * 1.5) * interpolate(frame, [25, 45], [6, 0])
      : 0;

  return (
    <AbsoluteFill
      style={{
        display: 'flex',
        flexDirection: 'row',
        alignItems: 'center',
        transform: `translateY(${shake}px)`,
      }}
    >
      <StageOverlay
        stageNum="3"
        tag="Stage 3 • Debt Avalanche"
        tagColor={COLORS.red}
        title="Crush 21% Cards First"
        rule="Rank debts strictly by interest rate. Crush high-rate credit cards and loans with full force. Leave low-rate index debt (HECS) calm."
        frame={frame}
        fps={fps}
      />
      <div style={{ flex: 1, display: 'flex', justifyContent: 'flex-end', paddingRight: 60, marginTop: 40 }}>
        <Img
          src={staticFile('illustrations/03-debt-avalanche-crusher.png')}
          style={{
            width: '72%',
            height: 'auto',
          }}
        />
      </div>
    </AbsoluteFill>
  );
};

const SceneStepFour = ({ frame, fps }) => {
  const scale = interpolate(frame, [0, 150], [1.02, 1.0], { extrapolateRight: 'clamp' });

  return (
    <AbsoluteFill style={{ display: 'flex', flexDirection: 'row', alignItems: 'center' }}>
      <StageOverlay
        stageNum="4"
        tag="Stage 4 • Compounding"
        tagColor={COLORS.accent}
        title="Stress-Tested Wealth Growth"
        rule="Survive variable rate spikes (+2.0%) under a defensive shield while smoothly spinning the gears of superannuation and compounding growth."
        frame={frame}
        fps={fps}
      />
      <div style={{ flex: 1, display: 'flex', justifyContent: 'flex-end', paddingRight: 60, marginTop: 40 }}>
        <Img
          src={staticFile('illustrations/04-stress-test-compounding.png')}
          style={{
            width: '72%',
            height: 'auto',
            transform: `scale(${scale})`,
            transformOrigin: 'center center',
          }}
        />
      </div>
    </AbsoluteFill>
  );
};

const SceneOutro = ({ frame, fps }) => {
  const cardSpring = spring({ frame, fps, config: { damping: 14 } });
  const cardScale = interpolate(cardSpring, [0, 1], [0.85, 1.0]);
  const cardOpacity = interpolate(cardSpring, [0, 1], [0, 1]);

  return (
    <AbsoluteFill style={{ justifyContent: 'center', alignItems: 'center', backgroundColor: '#F8FAFC' }}>
      <div
        style={{
          width: 820,
          padding: '48px 56px',
          backgroundColor: COLORS.white,
          border: `3px solid ${COLORS.navy}`,
          borderRadius: 24,
          boxShadow: '0 30px 60px rgba(10, 25, 47, 0.12)',
          textAlign: 'center',
          opacity: cardOpacity,
          transform: `scale(${cardScale})`,
          fontFamily: 'system-ui, -apple-system, sans-serif',
        }}
      >
        <div
          style={{
            display: 'inline-block',
            backgroundColor: COLORS.navy,
            color: COLORS.white,
            fontWeight: 800,
            fontSize: 14,
            padding: '6px 16px',
            borderRadius: 999,
            marginBottom: 16,
            letterSpacing: '0.05em',
          }}
        >
          MONEYMOVES AU DOSSIER
        </div>
        <h1
          style={{
            margin: '0 0 16px 0',
            fontSize: 40,
            fontWeight: 900,
            color: COLORS.navy,
            letterSpacing: '-0.03em',
          }}
        >
          Get Your Personalised Money Plan
        </h1>
        <p
          style={{
            margin: '0 0 28px 0',
            fontSize: 20,
            color: COLORS.inkSoft,
            lineHeight: 1.5,
          }}
        >
          A 12-month cash map, your debt payoff order, a 3-month income stress test and a dated action checklist, built from your own numbers.
        </p>

        <div
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: 20,
            padding: '12px 28px',
            backgroundColor: COLORS.cream,
            border: `1px solid ${COLORS.border}`,
            borderRadius: 14,
          }}
        >
          <span style={{ fontSize: 24, fontWeight: 900, color: COLORS.navy }}>Free to try</span>
          <span style={{ color: COLORS.inkSoft, fontSize: 16 }}>•</span>
          <span style={{ fontSize: 16, fontWeight: 700, color: COLORS.accent }}>moneymoves-au.vercel.app</span>
        </div>
      </div>
    </AbsoluteFill>
  );
};
