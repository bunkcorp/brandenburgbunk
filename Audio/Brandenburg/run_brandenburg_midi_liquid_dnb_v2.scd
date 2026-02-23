// Brandenburg MIDI – Liquid DnB v2 (corrected).
// Requires: Quarks.install("wslib")  // then recompile (Cmd+K) and run this (Cmd+Enter)
// Stop: ~stopBrandenburgMIDI.() or Cmd+Period.
//
// Fixes: Splay (no ! 2), proper 2-step snare (offset 1), sustain = dur*0.95,
//        filter sweep aligned to arrangement (intro/build/hold), sidechain locked to TempoClock,
//        reese → masterBus (dry, no verb on sub).

s.options.numInputBusChannels = 0;

s.waitForBoot {

  ~midiDir = "/Users/kevinwoods/Downloads/brand5-3-mids";
  ~partFiles = ["flute-score.mid", "cello-score.mid", "harpsichord-score.mid"];

  ~verbBus = Bus.audio(s, 2);
  ~masterBus = Bus.audio(s, 2);
  ~triggerBus = Bus.audio(s, 1);
  ~levelBus = Bus.control(s, 1);
  ~drumBus = Bus.audio(s, 2);

  // ----------------- REVERB + CHORUS + SWEEP (aligned to arrangement) -----------------
  // Intro 2s closed, build 8s opens, hold 90s open — no sweep at wrong time
  SynthDef(\lushVerb, { |verbBus = 0, masterBus = 0, intro = 2, build = 8, hold = 90|
    var dry, rate, mod, chorus, sweep, wet;
    dry = In.ar(verbBus, 2);
    rate = 0.2;
    mod = LFTri.kr(rate).range(0.005, 0.018);
    chorus = dry + (DelayC.ar(dry, 0.03, mod) * 0.28);
    sweep = EnvGen.kr(Env([400, 400, 6000, 6000], [intro, build, hold], \lin));
    chorus = LPF.ar(chorus, sweep);
    wet = FreeVerb2.ar(chorus[0], chorus[1], 0.5, 0.9, 0.3);
    Out.ar(masterBus, dry + (wet * 0.6));
  }).add;

  // ----------------- SIDECHAIN (locked to TempoClock) -----------------
  SynthDef(\sidechainTrigger, { |triggerBus = 0, tempo = 2|
    var trig = Impulse.kr(tempo);
    var env = EnvGen.ar(Env.perc(0.001, 0.12), trig);
    Out.ar(triggerBus, env);
  }).add;

  SynthDef(\sidechain, { |masterBus = 0, triggerBus = 0, levelBus = 0, drumBus = 0, duck = 0.5|
    var sig = In.ar(masterBus, 2);
    var sc = In.ar(triggerBus, 1);
    var gain = 1 - (sc * duck);
    var level = In.kr(levelBus, 1);
    var drums = In.ar(drumBus, 2);
    Out.ar(0, (sig * gain * level) + drums);
  }).add;

  SynthDef(\arrangementLevel, { |levelBus = 0, intro = 2, build = 8, hold = 90, breakdown = 10|
    var env = Env([0, 0, 1, 1, 0], [intro, build, hold, breakdown], \sin);
    var level = EnvGen.kr(env, doneAction: 2);
    Out.kr(levelBus, level);
  }).add;

  // ----------------- ORCHESTRA: Splay creates stereo (no ! 2) -----------------
  SynthDef(\brand_flute, { |freq = 440, amp = 0.3, sustain = 1, gate = 1, pan = 0, verbBus = 0|
    var env, sig, width;
    env = EnvGen.kr(Env.adsr(0.05, 0.3, 0.7, 0.4), gate, doneAction: 2);
    sig = Saw.ar(freq * [1, 1.003]).sum * 0.5;
    sig = LPF.ar(sig, 6000);
    width = SinOsc.kr(0.05).range(0.3, 0.8);
    Out.ar(verbBus, Splay.ar(sig * env * amp, width));
  }).add;

  SynthDef(\brand_cello, { |freq = 220, amp = 0.4, sustain = 1, gate = 1, pan = 0, verbBus = 0|
    var env, sig, width;
    env = EnvGen.kr(Env.adsr(0.1, 0.4, 0.8, 0.5), gate, doneAction: 2);
    sig = Saw.ar(freq * [1, 1.01]);
    sig = LPF.ar(sig.sum, 1200);
    width = SinOsc.kr(0.05).range(0.3, 0.8);
    Out.ar(verbBus, Splay.ar(sig * env * amp, width));
  }).add;

  SynthDef(\brand_harpsichord, { |freq = 440, amp = 0.25, sustain = 0.3, gate = 1, pan = 0, verbBus = 0|
    var env, sig, width;
    env = EnvGen.kr(Env.perc(0.003, sustain), doneAction: 2);
    sig = Pulse.ar(freq, 0.5);
    sig = HPF.ar(sig, 500);
    width = SinOsc.kr(0.05).range(0.3, 0.8);
    Out.ar(verbBus, Splay.ar(sig * env * amp, width));
  }).add;

  // ----------------- REESE → masterBus (dry, no verb on sub) -----------------
  SynthDef(\reeseBass, { |freq = 55, amp = 0.6, gate = 1, pan = 0, masterBus = 0|
    var env = EnvGen.kr(Env.adsr(0.01, 0.3, 0.8, 0.2), gate, doneAction: 2);
    var detune = [0.99, 1.01];
    var sig = Saw.ar(freq * detune);
    sig = LPF.ar(sig.sum, 800);
    sig = sig.tanh;
    Out.ar(masterBus, Pan2.ar(sig * env * amp, pan));
  }).add;

  SynthDef(\dnbKick, { |drumBus = 0|
    var env = EnvGen.ar(Env.perc(0.001, 0.2), doneAction: 2);
    var sig = SinOsc.ar(50, 0, env) * 0.8;
    Out.ar(drumBus, Pan2.ar(sig));
  }).add;

  SynthDef(\dnbSnare, { |drumBus = 0|
    var env = EnvGen.ar(Env.perc(0.005, 0.15), doneAction: 2);
    var noise = WhiteNoise.ar * env;
    Out.ar(drumBus, Pan2.ar(noise * 0.5));
  }).add;

  s.sync;

  TempoClock.default.tempo = 172 / 60;

  ~verb = Synth(\lushVerb, [
    \verbBus, ~verbBus.index,
    \masterBus, ~masterBus.index,
    \intro, 2, \build, 8, \hold, 90
  ], s, \addToTail);

  ~triggerSynth = Synth(\sidechainTrigger, [
    \triggerBus, ~triggerBus.index,
    \tempo, TempoClock.default.tempo
  ], s, \addToTail);

  ~sidechainSynth = Synth(\sidechain, [
    \masterBus, ~masterBus.index,
    \triggerBus, ~triggerBus.index,
    \levelBus, ~levelBus.index,
    \drumBus, ~drumBus.index,
    \duck, 0.4
  ], s, \addToTail);

  ~levelBus.set(1);
  ~testBeep = Synth(\brand_flute, [\freq, 523.25, \amp, 0.5, \pan, 0, \verbBus, ~verbBus.index], s, \addToHead);
  fork { 0.5.wait; ~testBeep.set(\gate, 0); 0.2.wait; ~levelBus.set(0) };
  "Brandenburg Liquid DnB v2: test beep (C5).".postln;

  ~patterns = [];
  ~partFiles.do { |name|
    var path = ~midiDir +/+ name;
    if (PathName(path).isFile) {
      var midi = SimpleMIDIFile.read(path);
      var pat = midi.p;
      ~patterns = ~patterns.add([name, pat]);
      ("Loaded: " ++ name).postln;
    } {
      ("Missing: " ++ name).postln;
    };
  };

  // ----------------- PLAYBACK: dur ×0.25, sustain = dur*0.95 (no double scaling) -----------------
  ~makePart = { |name, pat|
    var synth = \brand_flute, pan = 0;
    if (name.contains("cello")) { synth = \brand_cello; pan = -0.4 };
    if (name.contains("harpsichord")) { synth = \brand_harpsichord; pan = 0.4 };
    Pbindf(pat,
      \instrument, synth,
      \dur, Pkey(\dur) * 0.25,
      \sustain, Pkey(\dur) * 0.95,
      \amp, (Pkey(\amp) ? ((Pkey(\vel) ? 100) / 127)) * 0.5,
      \pan, pan,
      \verbBus, ~verbBus.index
    );
  };

  ~orchestraScore = Ppar(~patterns.collect { |pair| ~makePart.(pair[0], pair[1]) });

  ~bass = Pbind(
    \instrument, \reeseBass,
    \scale, Scale.minor,
    \degree, Pseq([0, -2, -3, -5], inf),
    \octave, 2,
    \dur, 0.5,
    \masterBus, ~masterBus.index
  );

  // Kick on 1 and 3; snare on 2 and 4 (first delta 1, then every 2 = backbeat)
  ~drums = Ppar([
    Pbind(\instrument, \dnbKick, \dur, 1, \drumBus, ~drumBus.index),
    Pbind(\instrument, \dnbSnare, \dur, Pseq([1, Pn(2, inf)], 1), \drumBus, ~drumBus.index)
  ]);

  ~score = Ppar([~orchestraScore, ~bass, ~drums]);

  ~startBrandenburg = {
    ~player = ~score.play;
    ~arrangementSynth = Synth(\arrangementLevel, [
      \levelBus, ~levelBus.index,
      \intro, 2, \build, 8, \hold, 90, \breakdown, 10
    ], s, \addToTail);
    "Brandenburg Liquid DnB v2 @ 172 BPM (corrected stereo, sweep, sustain, 2-step, reese dry).".postln;
    nil
  };
  ~startBrandenburg.();

  "v2: Splay fix, sustain=dur*0.95, sweep=arrangement, sidechain=tempo, snare offset 1, reese→master.".postln;

  ~stopBrandenburgMIDI = {
    ~player.stop;
    fork { 2.wait; s.freeAll; "Stopped (release tails cleared).".postln };
    nil
  };
  nil
};
~stopBrandenburgMIDI.();
