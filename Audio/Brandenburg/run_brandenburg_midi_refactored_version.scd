// Brandenburg MIDI – refactored version (SimpleMIDIFile + Patterns).
// Requires: Quarks.install("wslib")  // then recompile (Cmd+K) and run this (Cmd+Enter)
// Stop: ~stopBrandenburgMIDI.() or Cmd+Period.
//
// Architecture: send bus → chorus → reverb → master bus → sidechain + arrangement level → out.
// Cinematic: single shared reverb, slow chorus, light sidechain pump, intro/build/drop/breakdown.

s.options.numInputBusChannels = 0;

s.waitForBoot {

  ~midiDir = "/Users/kevinwoods/Downloads/brand5-3-mids";
  ~partFiles = ["flute-score.mid", "cello-score.mid", "harpsichord-score.mid"];

  // ----------------- BUSSES -----------------
  ~verbBus = Bus.audio(s, 2);
  ~masterBus = Bus.audio(s, 2);
  ~triggerBus = Bus.audio(s, 1);   // sidechain trigger (virtual kick)
  ~levelBus = Bus.control(s, 1);   // arrangement: 0 = silent, 1 = full

  // ----------------- SHARED REVERB + SLOW CHORUS (one instance) -----------------
  SynthDef(\lushVerb, { |verbBus = 0, masterBus = 0|
    var dry = In.ar(verbBus, 2);
    var rate = 0.2;
    var mod = LFTri.kr(rate).range(0.005, 0.018);
    var chorus = dry + (DelayC.ar(dry, 0.03, mod) * 0.28);
    var wet = FreeVerb2.ar(chorus[0], chorus[1], 0.5, 0.9, 0.3);
    Out.ar(masterBus, dry + (wet * 0.6));
  }).add;

  // ----------------- SIDECHAIN (duck mix on virtual kick) -----------------
  SynthDef(\sidechainTrigger, { |triggerBus = 0, bpm = 60|
    var trig = Impulse.kr(bpm / 60);
    var env = EnvGen.ar(Env.perc(0.001, 0.12), trig);
    Out.ar(triggerBus, env);
  }).add;

  SynthDef(\sidechain, { |masterBus = 0, triggerBus = 0, levelBus = 0, duck = 0.5|
    var sig = In.ar(masterBus, 2);
    var sc = In.ar(triggerBus, 1);
    var gain = 1 - (sc * duck);
    var level = In.kr(levelBus, 1);
    Out.ar(0, sig * gain * level);
  }).add;

  // ----------------- ARRANGEMENT ENVELOPE (intro → build → drop → breakdown) -----------------
  SynthDef(\arrangementLevel, { |levelBus = 0, intro = 2, build = 8, hold = 90, breakdown = 10|
    var env = Env([0, 0, 1, 1, 0], [intro, build, hold, breakdown], \sin);
    var level = EnvGen.kr(env, doneAction: 2);
    Out.kr(levelBus, level);
  }).add;

  // ----------------- INSTRUMENTS: dry to verb bus, Pan2 for width, ADSR gate (no per-note reverb) -----------------

  SynthDef(\brand_flute, { |freq = 440, amp = 0.3, sustain = 1, gate = 1, pan = 0, verbBus = 0|
    var env = EnvGen.kr(Env.adsr(0.05, 0.3, 0.7, 0.4), gate, doneAction: 2);
    var sig = Saw.ar(freq * [1, 1.003]).sum * 0.5;
    sig = LPF.ar(sig, 6000);
    Out.ar(verbBus, Pan2.ar(sig * env * amp, pan));
  }).add;

  SynthDef(\brand_cello, { |freq = 220, amp = 0.4, sustain = 1, gate = 1, pan = 0, verbBus = 0|
    var env = EnvGen.kr(Env.adsr(0.1, 0.4, 0.8, 0.5), gate, doneAction: 2);
    var sig = Saw.ar(freq * [1, 1.01]);
    sig = LPF.ar(sig.sum, 1200);
    Out.ar(verbBus, Pan2.ar(sig * env * amp, pan));
  }).add;

  SynthDef(\brand_harpsichord, { |freq = 440, amp = 0.25, sustain = 0.3, gate = 1, pan = 0, verbBus = 0|
    var env = EnvGen.kr(Env.perc(0.003, sustain), doneAction: 2);
    var sig = Pulse.ar(freq, 0.5);
    sig = HPF.ar(sig, 500);
    Out.ar(verbBus, Pan2.ar(sig * env * amp, pan));
  }).add;

  s.sync;

  // Effect chain: verb (chorus+reverb) → masterBus → sidechain (duck + level) → 0
  ~verb = Synth(\lushVerb, [\verbBus, ~verbBus.index, \masterBus, ~masterBus.index], s, \addToTail);
  ~triggerSynth = Synth(\sidechainTrigger, [\triggerBus, ~triggerBus.index, \bpm, 60], s, \addToTail);
  ~sidechainSynth = Synth(\sidechain, [\masterBus, ~masterBus.index, \triggerBus, ~triggerBus.index, \levelBus, ~levelBus.index, \duck, 0.4], s, \addToTail);

  // Test beep (level 1 so you hear it; then level 0 for arrangement fade-in)
  ~levelBus.set(1);
  ~testBeep = Synth(\brand_flute, [\freq, 523.25, \amp, 0.5, \pan, 0, \verbBus, ~verbBus.index], s, \addToHead);
  fork { 0.5.wait; ~testBeep.set(\gate, 0); 0.2.wait; ~levelBus.set(0) };
  "Brandenburg: test beep (C5). If no sound, set Server → Output Device.".postln;

  // ----------------- LOAD MIDI (SimpleMIDIFile – no custom parser) -----------------

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

  // ----------------- BUILD PLAYBACK (Ppar = all parts in parallel) -----------------

  ~makePart = { |name, pat|
    var synth = \brand_flute, pan = 0;
    if (name.contains("cello")) { synth = \brand_cello; pan = -0.4 };
    if (name.contains("harpsichord")) { synth = \brand_harpsichord; pan = 0.4 };
    Pbindf(pat,
      \instrument, synth,
      \amp, (Pkey(\amp) ? ((Pkey(\vel) ? 100) / 127)) * 0.5,
      \pan, pan,
      \verbBus, ~verbBus.index,
      \sustain, (Pkey(\sustain) ? (Pkey(\dur) * 0.95)) * 0.95
    );
  };

  ~score = Ppar(~patterns.collect { |pair| ~makePart.(pair[0], pair[1]) });

  ~startBrandenburg = {
    ~player = ~score.play;
    ~arrangementSynth = Synth(\arrangementLevel, [
      \levelBus, ~levelBus.index,
      \intro, 2,   // 2 s silence (intro)
      \build, 8,   // 8 s fade in
      \hold, 90,   // full level for 90 s (drop)
      \breakdown, 10  // 10 s fade out
    ], s, \addToTail);
    "Brandenburg started (intro → build → drop → breakdown).".postln;
    nil
  };
  ~startBrandenburg.();

  "Sidechain: 60 BPM pump; chorus on verb; arrangement: 2s intro, 8s build, 90s hold, 10s out.".postln;

  // Optional liquid DnB: keep lush pad at original tempo, then e.g.:
  // TempoClock.default.tempo = 172/60;
  // In ~makePart add: \dur, Pkey(\dur) * 0.25
  // Plus: reese bass, 2-step drums, light sidechain for a proper drop.

  ~stopBrandenburgMIDI = {
    ~player.stop;
    fork { 2.wait; s.freeAll; "Stopped (release tails cleared).".postln };
    nil
  };
  // Restart (new arrangement): ~startBrandenburg.()
  nil  // so post window shows -> nil instead of -> a Routine
};
// ~stopBrandenburgMIDI.(); 