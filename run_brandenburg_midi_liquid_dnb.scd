// Brandenburg MIDI – Liquid DnB version.
// Requires: Quarks.install("wslib")  // then recompile (Cmd+K) and run this (Cmd+Enter)
// Stop: ~stopBrandenburgMIDI.() or Cmd+Period.
//
// Same architecture + Liquid DnB upgrades:
// • 172 BPM, compressed note lengths (×0.25)
// • Reese bass (through verb)
// • Kick + snare (dry, mixed at output)
// • Filter sweep on build (lushVerb opens 400→6k over 10s)
// • Evolving stereo width (Splay + slow LFO) on orchestra

s.options.numInputBusChannels = 0;

s.waitForBoot {

  ~midiDir = "/Users/kevinwoods/Downloads/brand5-3-mids";
  ~partFiles = ["flute-score.mid", "cello-score.mid", "harpsichord-score.mid"];

  // ----------------- BUSSES -----------------
  ~verbBus = Bus.audio(s, 2);
  ~masterBus = Bus.audio(s, 2);
  ~triggerBus = Bus.audio(s, 1);
  ~levelBus = Bus.control(s, 1);
  ~drumBus = Bus.audio(s, 2);   // kick/snare (dry, mixed at output)

  // ----------------- SHARED REVERB + CHORUS + BUILD SWEEP -----------------
  SynthDef(\lushVerb, { |verbBus = 0, masterBus = 0|
    var dry, rate, mod, chorus, sweep, wet;
    dry = In.ar(verbBus, 2);
    rate = 0.2;
    mod = LFTri.kr(rate).range(0.005, 0.018);
    chorus = dry + (DelayC.ar(dry, 0.03, mod) * 0.28);
    sweep = XLine.kr(400, 6000, 10);  // build: filter opens over 10 s
    chorus = LPF.ar(chorus, sweep);
    wet = FreeVerb2.ar(chorus[0], chorus[1], 0.5, 0.9, 0.3);
    Out.ar(masterBus, dry + (wet * 0.6));
  }).add;

  // ----------------- SIDECHAIN (172 BPM liquid pump) -----------------
  SynthDef(\sidechainTrigger, { |triggerBus = 0, bpm = 172|
    var trig = Impulse.kr(bpm / 60);
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

  // ----------------- ORCHESTRA: evolving stereo width (Splay + LFO) -----------------
  SynthDef(\brand_flute, { |freq = 440, amp = 0.3, sustain = 1, gate = 1, pan = 0, verbBus = 0|
    var env, sig, width;
    env = EnvGen.kr(Env.adsr(0.05, 0.3, 0.7, 0.4), gate, doneAction: 2);
    sig = Saw.ar(freq * [1, 1.003]).sum * 0.5;
    sig = LPF.ar(sig, 6000);
    width = SinOsc.kr(0.05).range(0.3, 0.8);
    Out.ar(verbBus, Splay.ar((sig * env * amp) ! 2, width));
  }).add;

  SynthDef(\brand_cello, { |freq = 220, amp = 0.4, sustain = 1, gate = 1, pan = 0, verbBus = 0|
    var env, sig, width;
    env = EnvGen.kr(Env.adsr(0.1, 0.4, 0.8, 0.5), gate, doneAction: 2);
    sig = Saw.ar(freq * [1, 1.01]);
    sig = LPF.ar(sig.sum, 1200);
    width = SinOsc.kr(0.05).range(0.3, 0.8);
    Out.ar(verbBus, Splay.ar((sig * env * amp) ! 2, width));
  }).add;

  SynthDef(\brand_harpsichord, { |freq = 440, amp = 0.25, sustain = 0.3, gate = 1, pan = 0, verbBus = 0|
    var env, sig, width;
    env = EnvGen.kr(Env.perc(0.003, sustain), doneAction: 2);
    sig = Pulse.ar(freq, 0.5);
    sig = HPF.ar(sig, 500);
    width = SinOsc.kr(0.05).range(0.3, 0.8);
    Out.ar(verbBus, Splay.ar((sig * env * amp) ! 2, width));
  }).add;

  // ----------------- REESE BASS (through verb) -----------------
  SynthDef(\reeseBass, { |freq = 55, amp = 0.6, gate = 1, pan = 0, verbBus = 0|
    var env = EnvGen.kr(Env.adsr(0.01, 0.3, 0.8, 0.2), gate, doneAction: 2);
    var detune = [0.99, 1.01];
    var sig = Saw.ar(freq * detune);
    sig = LPF.ar(sig.sum, 800);
    sig = sig.tanh;
    Out.ar(verbBus, Pan2.ar(sig * env * amp, pan));
  }).add;

  // ----------------- LIQUID KICK + SNARE (dry → drum bus) -----------------
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

  ~verb = Synth(\lushVerb, [\verbBus, ~verbBus.index, \masterBus, ~masterBus.index], s, \addToTail);
  ~triggerSynth = Synth(\sidechainTrigger, [\triggerBus, ~triggerBus.index, \bpm, 172], s, \addToTail);
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
  "Brandenburg Liquid DnB: test beep (C5).".postln;

  // ----------------- LOAD MIDI -----------------
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

  // ----------------- BUILD PLAYBACK (compressed durs ×0.25 for 172 BPM) -----------------
  ~makePart = { |name, pat|
    var synth = \brand_flute, pan = 0;
    if (name.contains("cello")) { synth = \brand_cello; pan = -0.4 };
    if (name.contains("harpsichord")) { synth = \brand_harpsichord; pan = 0.4 };
    Pbindf(pat,
      \instrument, synth,
      \dur, Pkey(\dur) * 0.25,
      \amp, (Pkey(\amp) ? ((Pkey(\vel) ? 100) / 127)) * 0.5,
      \pan, pan,
      \verbBus, ~verbBus.index,
      \sustain, (Pkey(\sustain) ? (Pkey(\dur) * 0.95)) * 0.95
    );
  };

  ~orchestraScore = Ppar(~patterns.collect { |pair| ~makePart.(pair[0], pair[1]) });

  ~bass = Pbind(
    \instrument, \reeseBass,
    \scale, Scale.minor,
    \degree, Pseq([0, -2, -3, -5], inf),
    \octave, 2,
    \dur, 0.5,
    \verbBus, ~verbBus.index
  );

  ~drums = Ppar([
    Pbind(\instrument, \dnbKick, \dur, 1, \drumBus, ~drumBus.index),
    Pbind(\instrument, \dnbSnare, \dur, Pseq([1, Pn(2, inf)], 1), \drumBus, ~drumBus.index)
  ]);

  ~score = Ppar([~orchestraScore, ~bass, ~drums]);

  ~startBrandenburg = {
    ~player = ~score.play;
    ~arrangementSynth = Synth(\arrangementLevel, [
      \levelBus, ~levelBus.index,
      \intro, 2,
      \build, 8,
      \hold, 90,
      \breakdown, 10
    ], s, \addToTail);
    "Brandenburg Liquid DnB started @ 172 BPM (orchestra + reese + kick/snare).".postln;
    nil
  };
  ~startBrandenburg.();

  "Liquid DnB: 172 BPM, filter sweep build, evolving width, reese + 2-step drums.".postln;

  ~stopBrandenburgMIDI = {
    ~player.stop;
    fork { 2.wait; s.freeAll; "Stopped (release tails cleared).".postln };
    nil
  };
  nil
};
~stopBrandenburgMIDI.();
