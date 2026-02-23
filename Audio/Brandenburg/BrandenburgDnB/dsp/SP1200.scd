/*
  SP-1200 Emulation Module
  E-mu SP-1200: 12-bit, 26.04 kHz, aliasing, saturation, analog noise.
  Low CPU: single-pass downsample → quantize → soft clip → noise.
*/
(
SynthDef(\sp1200_mono, { |inBus, outBus = 0, bitDepth = 12, sampleRate = 26040, drive = 0.25, noise = 0.02, gate = 1|
  var snd = In.ar(inBus, 1);
  var sr = SampleRate.ir;
  var ratio = sampleRate / sr;
  // Downsample: Latch at ~26k (sample-and-hold style)
  snd = Latch.ar(snd, Impulse.ar(sampleRate));
  // 12-bit quantization: 2^12 levels, map -1..1 to 0..4095 then back
  snd = (snd * (2 ** bitDepth)).round / (2 ** bitDepth);
  // Mild soft clipping (saturation)
  snd = (snd * (1 + (drive * 4))).tanh;
  // Subtle analog noise floor
  snd = snd + (WhiteNoise.ar(noise) * EnvGen.kr(Env.asr(0.01, 1, 0.01), gate, doneAction: 0));
  Out.ar(outBus, snd ! 2);
}).add;

SynthDef(\sp1200_stereo, { |inBus, outBus = 0, bitDepth = 12, sampleRate = 26040, drive = 0.25, noise = 0.02, gate = 1|
  var snd = In.ar(inBus, 2);
  var sr = SampleRate.ir;
  var rate = sampleRate / sr;
  snd = Latch.ar(snd, Impulse.ar(sampleRate, 0, 1, 0));
  snd = (snd * (2 ** bitDepth)).round / (2 ** bitDepth);
  snd = (snd * (1 + (drive * 4))).tanh;
  snd = snd + (WhiteNoise.ar(noise ! 2) * EnvGen.kr(Env.asr(0.01, 1, 0.01), gate, doneAction: 0));
  Out.ar(outBus, snd);
}).add;
);
