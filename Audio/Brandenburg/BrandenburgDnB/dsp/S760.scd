/*
  S-760 Clean Mode
  Roland S-760: 16-bit, 44.1 kHz, clean digital, optional 12dB lowpass, subtle chorus.
*/
(
SynthDef(\s760_stereo, { |inBus, outBus = 0, cutoff = 20000, chorusRate = 0.5, chorusDepth = 0.003, gate = 1|
  var snd = In.ar(inBus, 2);
  snd = LPF.ar(snd, cutoff.clip(20, 20000));
  // Subtle chorus: slow LFO on short delay
  snd = snd + DelayC.ar(snd, 0.02, SinOsc.kr(chorusRate, 0).range(0.003, 0.007)) * chorusDepth;
  snd = snd * EnvGen.kr(Env.asr(0.01, 1, 0.01), gate, doneAction: 0);
  Out.ar(outBus, snd);
}).add;
);
