/*
  Ethereal FX Chain
  Order: Highpass 120Hz → Hall reverb → Ping-pong delay → Chorus → (sidechain in main)
  Reverb: large hall, decay 4.5–7.5s, pre-delay 20ms, wet 20–35%, highcut 6k.
  Delay: ping-pong, 1/8 dotted, feedback 35–45%, lowcut 400Hz, wet 15–25%.
*/
(
SynthDef(\fx_hp_reverb, { |inBus, outBus = 0, hpFreq = 120, revDecay = 5.5, revPredelay = 0.02, revWet = 0.28, revHighCut = 6000, gate = 1|
  var snd = In.ar(inBus, 2);
  snd = HPF.ar(snd, hpFreq);
  snd = snd + GVerb.ar(snd.sum, 80, revDecay, 0.5, 0.5, 15, revPredelay, 0.7, 0.5);
  snd = LPF.ar(snd, revHighCut);
  snd = (snd * (1 - revWet)) + (GVerb.ar(snd.sum, 100, revDecay, 0.5, 0.5, 15, revPredelay, 0.7, 0.5) * revWet);
  snd = snd * EnvGen.kr(Env.asr(0.01, 1, 0.01), gate, doneAction: 0);
  Out.ar(outBus, snd);
}).add;

// Simpler single reverb (GVerb is stereo in, stereo out in recent SC)
SynthDef(\fx_ethereal_reverb, { |inBus, outBus = 0, decay = 5.5, predelay = 0.02, wet = 0.28, highCut = 6000, gate = 1|
  var dry = In.ar(inBus, 2);
  var wetSig = GVerb.ar(dry.sum, 100, decay, 0.5, 0.5, 15, predelay, 0.7, 0.5);
  wetSig = LPF.ar(wetSig, highCut);
  dry = HPF.ar(dry, 120);
  Out.ar(outBus, (dry * (1 - wet)) + (wetSig * wet) * EnvGen.kr(Env.asr(0.01, 1, 0.01), gate, doneAction: 0));
}).add;

SynthDef(\fx_pingpong_delay, { |inBus, outBus = 0, time = 0.272, feedback = 0.4, lowCut = 400, wet = 0.2, gate = 1|
  var snd = In.ar(inBus, 2);
  var left = DelayC.ar(snd[0], 1, time);
  var right = DelayC.ar(snd[1], 1, time);
  var crossL = LPF.ar(right * feedback, 8000);
  var crossR = LPF.ar(left * feedback, 8000);
  left = LPF.ar(left + crossL, 8000);
  right = LPF.ar(right + crossR, 8000);
  snd = (snd * (1 - wet)) + (HPF.ar([left, right], lowCut) * wet);
  Out.ar(outBus, snd * EnvGen.kr(Env.asr(0.01, 1, 0.01), gate, doneAction: 0));
}).add;

SynthDef(\fx_chorus, { |inBus, outBus = 0, rate = 0.3, depth = 0.004, gate = 1|
  var snd = In.ar(inBus, 2);
  snd = snd + DelayC.ar(snd, 0.05, SinOsc.kr(rate, [0, 1]).range(0.01, 0.03)) * depth;
  Out.ar(outBus, snd * EnvGen.kr(Env.asr(0.01, 1, 0.01), gate, doneAction: 0));
}).add;

// Sidechain: duck inBus by trigger (e.g. kick). Attack fast, release ~0.2–0.4s.
SynthDef(\fx_sidechain, { |inBus, outBus = 0, triggerBus = 0, attack = 0.002, release = 0.25, threshold = 0.3, gate = 1|
  var snd = In.ar(inBus, 2);
  var trig = In.kr(triggerBus, 1);
  var env = EnvGen.kr(Env.perc(attack, release, 1, -4), trig);
  snd = snd * (1 - (env * (1 - threshold)));
  Out.ar(outBus, snd * EnvGen.kr(Env.asr(0.01, 1, 0.01), gate, doneAction: 0));
}).add;
);
