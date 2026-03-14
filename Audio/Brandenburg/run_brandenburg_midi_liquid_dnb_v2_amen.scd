// Brandenburg MIDI – Liquid DnB v2 (Amen version).
// Requires: Quarks.install("wslib")  // then recompile (Cmd+K) and run this (Cmd+Enter)
// Stop: ~stopBrandenburgMIDI.() or Cmd+Period.
//
// Changes vs v2:
// - Replaces synthetic kick/snare with Amen sample playback
// - Uses sample-accurate playback rate + bar retrigger sync at 172 BPM
// - Keeps v2 orchestral/reese/sidechain architecture

s.options.numInputBusChannels = 0;
s.options.blockSize = 512;
s.options.outDevice = "MacBook Pro Speakers";

s.waitForBoot {

  ~midiDir = "/Users/kevinwoods/Downloads/brand5-3-mids";
  ~partFiles = ["flute-score.mid", "cello-score.mid", "harpsichord-score.mid"];
  ~amenPath = "/Users/kevinwoods/Desktop/ActuarialExams/Audio/copied_patches/amen.wav";
  ~amenBeats = 4;
  ~amenBpm = 172;

  ~verbBus = Bus.audio(s, 2);
  ~masterBus = Bus.audio(s, 2);
  ~triggerBus = Bus.audio(s, 1);
  ~levelBus = Bus.control(s, 1);
  ~drumBus = Bus.audio(s, 2);

  // ----------------- REVERB + CHORUS + SWEEP (aligned to arrangement) -----------------
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
    // Trigger faster than quarter-note pulse so pumping follows break transients more tightly.
    var trig = Impulse.kr(tempo * 2);
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

  SynthDef(\arrangementLevel, { |levelBus = 0, intro = 0, build = 0.05, hold = 120, breakdown = 10|
    var env = Env([0, 0, 1, 1, 0], [intro, build, hold, breakdown], \sin);
    var level = EnvGen.kr(env, doneAction: 2);
    Out.kr(levelBus, level);
  }).add;

  // ----------------- ORCHESTRA: Splay creates stereo -----------------
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

  // ----------------- REESE -----------------
  SynthDef(\reeseBass, { |freq = 55, amp = 0.6, gate = 1, pan = 0, masterBus = 0|
    var env = EnvGen.kr(Env.adsr(0.01, 0.3, 0.8, 0.2), gate, doneAction: 2);
    var detune = [0.99, 1.01];
    var sig = Saw.ar(freq * detune);
    sig = LPF.ar(sig.sum, 800);
    sig = sig.tanh;
    Out.ar(masterBus, Pan2.ar(sig * env * amp, pan));
  }).add;

  // ----------------- AMEN DRUM LOOP -----------------
  SynthDef(\amenBreak, { |drumBus = 0, buf = 0, rate = 1, amp = 0.34, barRate = 0.71667|
    var trig = Impulse.kr(barRate);
    var sig = PlayBuf.ar(1, buf, rate, trigger: trig, startPos: 0);
    // Shape toward DnB break feel without fully replacing the source character
    sig = HPF.ar(sig, 30);
    sig = BPeakEQ.ar(sig, 1800, 0.8, 2.0);
    sig = (sig * 1.2).tanh;
    Out.ar(drumBus, Pan2.ar(sig * amp));
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
  fork { 0.5.wait; ~testBeep.set(\gate, 0) };
  "Brandenburg Liquid DnB v2 Amen: test beep (C5).".postln;

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

  ~makePart = { |name, pat|
    var synth = \brand_flute, pan = 0;
    if (name.contains("cello")) { synth = \brand_cello; pan = -0.4 };
    if (name.contains("harpsichord")) { synth = \brand_harpsichord; pan = 0.4 };
    Pbindf(pat,
      \instrument, synth,
      // Keep MIDI durations in beat units from SimpleMIDIFile.
      \dur, Pkey(\dur),
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

  // Amen loop sync math: derive loop duration from the actual loaded file.
  ~barSec = 4 * 60 / ~amenBpm;
  ~drums = Pbind(
    \instrument, \amenBreak,
    \dur, 9999,
    \drumBus, ~drumBus.index,
    \buf, Pfunc { ~amenBuf.bufnum },
    \rate, Pfunc {
      var loopSec = ~amenBuf.numFrames / s.sampleRate;
      loopSec / ~barSec
    },
    \barRate, 1 / ~barSec,
    \amp, 0.5
  );

  ~score = Ppar([~orchestraScore, ~bass, ~drums]);

  ~startBrandenburg = {
    ~player = ~score.play(quant: 4);
    ~arrangementSynth = Synth(\arrangementLevel, [
      \levelBus, ~levelBus.index,
      \intro, 0, \build, 0.05, \hold, 120, \breakdown, 10
    ], s, \addToTail);
    "Brandenburg Liquid DnB v2 Amen @ 172 BPM.".postln;
    nil
  };

  ~stopBrandenburgMIDI = {
    ~player.stop;
    fork { 2.wait; s.freeAll; "Stopped (release tails cleared).".postln };
    nil
  };

  ~amenBuf = Buffer.read(s, ~amenPath, action: {
    "Amen loaded.".postln;
    ~startBrandenburg.();
  });

  "v2-amen: Amen break replaces synthetic kick/snare; bar-synced retrigger enabled.".postln;
  nil
};
// ~stopBrandenburgMIDI.();
