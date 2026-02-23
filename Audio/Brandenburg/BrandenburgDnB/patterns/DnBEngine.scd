/*
  Drum & Bass Engine – 170–174 BPM, Amen-style / 2-step.
  Kick (tight), snare (sharp), ghost snares, optional reversed cymbal.
  Uses sample playback or simple synth kicks/snares; BPM and swing (53–56%) configurable.
*/
(
// Default BPM for DnB
~dnbBpm = 172;
~dnbSwing = 0.54; // 54% swing

// One bar at 172 BPM in seconds
~dnbBarSec = { 60 / ~dnbBpm * 4 };

// Tempo clock for scheduling (optional; can also drive from Tidal)
~dnbClock = TempoClock(~dnbBpm / 60, 0);

SynthDef(\dnb_kick, { |outBus = 0, amp = 0.9, decay = 0.25, freq = 55, gate = 1|
  var env = EnvGen.kr(Env.perc(0.005, decay, amp, -4), gate, doneAction: 2);
  var snd = SinOsc.ar(freq * EnvGen.kr(Env.perc(0.001, 0.1, 1, -8))) * env;
  snd = snd + (LPF.ar(WhiteNoise.ar(0.4), 800) * env);
  Out.ar(outBus, snd ! 2);
}).add;

SynthDef(\dnb_snare, { |outBus = 0, amp = 0.7, decay = 0.15, gate = 1|
  var env = EnvGen.kr(Env.perc(0.001, decay, amp, -4), gate, doneAction: 2);
  var noise = BPF.ar(WhiteNoise.ar(1), 1000, 0.5);
  var tone = SinOsc.ar(180 * EnvGen.kr(Env.perc(0.001, 0.05, 1, -8)));
  Out.ar(outBus, (noise + tone) * env ! 2);
}).add;

SynthDef(\dnb_ghost_snare, { |outBus = 0, amp = 0.25, decay = 0.08, gate = 1|
  var env = EnvGen.kr(Env.perc(0.001, decay, amp, -4), gate, doneAction: 2);
  var noise = BPF.ar(WhiteNoise.ar(0.6), 1200, 0.6);
  Out.ar(outBus, noise * env ! 2);
}).add;

SynthDef(\dnb_hat, { |outBus = 0, amp = 0.2, decay = 0.05, gate = 1|
  var env = EnvGen.kr(Env.perc(0.001, decay, amp, -4), gate, doneAction: 2);
  var snd = HPF.ar(WhiteNoise.ar(1), 7000) * env;
  Out.ar(outBus, snd ! 2);
}).add;
);
