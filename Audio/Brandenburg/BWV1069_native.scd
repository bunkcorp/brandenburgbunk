/*
  BWV 1069 – Orchestral Suite No. 4 in D major, Ouverture
  Simplified rendition using only native SuperCollider UGens (no samples).
  Run: s.boot; then run this file. Stop: Cmd+Period or ~stopBWV1069.()
*/
(
// D major: D=0, E=2, F#=4, G=5, A=7, B=9, C#=11. midi = 62 + degree
var d = 62; // D4
var e = 64;
var fs = 66;
var g = 67;
var a = 69;
var b = 71;
var cs5 = 73;
var d5 = 74;
var e5 = 76;
var fs5 = 78;
var g5 = 79;
var a5 = 81;

// Native instrument SynthDefs (no samples)
SynthDef(\bwv_trumpet, { |freq = 440, amp = 0.25, sustain = 0.4, gate = 1|
  var env = EnvGen.kr(Env.linen(0.02, sustain, 0.05, amp, -4), gate, doneAction: 2);
  var snd = Saw.ar(freq) + Saw.ar(freq * 1.5) * 0.5;
  snd = LPF.ar(snd, 1200);
  snd = snd + (SinOsc.ar(freq) * 0.4);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\bwv_string, { |freq = 440, amp = 0.2, sustain = 0.5, gate = 1|
  var env = EnvGen.kr(Env.linen(0.05, sustain, 0.1, amp, -4), gate, doneAction: 2);
  var snd = Saw.ar(freq) + Saw.ar(freq * 2) * 0.3;
  snd = LPF.ar(snd, 2000);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\bwv_timpani, { |freq = 80, amp = 0.4, gate = 1|
  var env = EnvGen.kr(Env.perc(0.01, 0.6, amp, -4), gate, doneAction: 2);
  var snd = SinOsc.ar(freq * EnvGen.kr(Env.perc(0.001, 0.2, 1, -8))) + LPF.ar(WhiteNoise.ar(0.3), 400);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\bwv_oboe, { |freq = 440, amp = 0.22, sustain = 0.35, gate = 1|
  var env = EnvGen.kr(Env.linen(0.02, sustain, 0.05, amp, -4), gate, doneAction: 2);
  var snd = Saw.ar(freq) + Pulse.ar(freq, 0.3) * 0.4;
  snd = BPF.ar(snd, freq * 2, 0.3);
  Out.ar(0, snd * env ! 2);
}).add;

s.sync;

// Ensure root group exists (server can have empty tree right after boot)
s.sendMsg("/g_new", 1, 0, 0);
s.sync;
~bwvGroup = 1;  // target all synths to root

// French Overture: Grave (slow dotted) then Allegro. BPM approx: Grave 60, Allegro 120.
~bpmGrave = 52;
~bpmAllegro = 116;
~tGrave = 60 / ~bpmGrave;
~tAllegro = 60 / ~bpmAllegro;

// Opening Grave: dotted rhythm fanfare in D major (simplified)
~playGrave = {
  var t = ~tGrave;
  var dot = t * 0.75;   // dotted quarter
  var eighth = t * 0.25;
  Routine({
    // Bar 1–2: D - A - D (fanfare)
    Synth(\bwv_trumpet, [\freq, d.midicps, \sustain, dot], ~bwvGroup); (dot).wait;
    Synth(\bwv_trumpet, [\freq, a.midicps, \sustain, eighth], ~bwvGroup); (eighth).wait;
    Synth(\bwv_trumpet, [\freq, d5.midicps, \sustain, dot], ~bwvGroup); (dot).wait;
    Synth(\bwv_timpani, [\freq, 73.4], ~bwvGroup); 0.1.wait;
    Synth(\bwv_string, [\freq, d.midicps, \sustain, t * 2], ~bwvGroup); (t).wait;
    // Bar 3–4: F# - A - D
    Synth(\bwv_trumpet, [\freq, fs.midicps, \sustain, dot], ~bwvGroup); (dot).wait;
    Synth(\bwv_trumpet, [\freq, a.midicps, \sustain, eighth], ~bwvGroup); (eighth).wait;
    Synth(\bwv_trumpet, [\freq, d5.midicps, \sustain, dot], ~bwvGroup); (dot).wait;
    Synth(\bwv_string, [\freq, fs.midicps, \sustain, t], ~bwvGroup); (t).wait;
    // Bar 5–6: scale down D - C# - B - A
    [d5, cs5, b, a].do { |midi|
      Synth(\bwv_oboe, [\freq, midi.midicps, \sustain, dot * 0.6], ~bwvGroup); (dot * 0.5).wait;
    };
    (t).wait;
    // Bar 7–8: G - B - D
    Synth(\bwv_trumpet, [\freq, g.midicps, \sustain, dot], ~bwvGroup); (dot).wait;
    Synth(\bwv_trumpet, [\freq, b.midicps, \sustain, eighth], ~bwvGroup); (eighth).wait;
    Synth(\bwv_trumpet, [\freq, d5.midicps, \sustain, dot], ~bwvGroup); (dot).wait;
    Synth(\bwv_timpani, [\freq, 73.4], ~bwvGroup); 0.1.wait;
    Synth(\bwv_string, [\freq, g.midicps, \sustain, t], ~bwvGroup); (t).wait;
    // Repeat opening for a bit then go to allegro
    2.do {
      Synth(\bwv_trumpet, [\freq, d.midicps, \sustain, dot], ~bwvGroup); (dot).wait;
      Synth(\bwv_trumpet, [\freq, a.midicps, \sustain, eighth], ~bwvGroup); (eighth).wait;
      Synth(\bwv_trumpet, [\freq, d5.midicps, \sustain, dot], ~bwvGroup); (dot).wait;
      (t).wait;
    };
  }).play;
};

// Allegro: fugal subject style in D major (quarter = ~tAllegro)
~playAllegro = {
  var t = ~tAllegro;
  var subj = [d5, a, d5, fs5, a5, g5, fs5, e5, d5, a, d5, fs5, a5, g5, fs5, e5]; // simplified subject
  Routine({
    subj.do { |midi, i|
      Synth(\bwv_string, [\freq, midi.midicps, \sustain, t * 0.9], ~bwvGroup);
      if (i % 4 == 0) { Synth(\bwv_timpani, [\freq, 73.4], ~bwvGroup) };
      (t).wait;
    };
    // Second entry (dominant)
    [a5, e5, a5, cs5, e5, d5, cs5, b, a5, e5, a5, cs5, e5, d5, cs5, b].do { |midi, i|
      Synth(\bwv_oboe, [\freq, midi.midicps, \sustain, t * 0.85], ~bwvGroup);
      (t).wait;
    };
    // Continuation
    8.do { |i|
      var n = [d5, fs5, a5, d5, fs5, g5, a5, fs5].wrapAt(i);
      Synth(\bwv_string, [\freq, n.midicps, \sustain, t * 0.8], ~bwvGroup); (t).wait;
    };
  }).play;
};

// Full Ouverture: Grave then Allegro
~runBWV1069 = {
  ~stopBWV1069.();
  ~bwv1069Routine = Routine({
    "BWV 1069 Ouverture – Grave".postln;
    ~playGrave.();
    (12 * ~tGrave).wait; // let grave play
    "BWV 1069 – Allegro".postln;
    ~playAllegro.();
    (20 * ~tAllegro).wait;
    "BWV 1069 – end".postln;
  }).play;
};
~stopBWV1069 = { ~bwv1069Routine !? { ~bwv1069Routine.stop }; s.freeAll; };

"BWV 1069 native loaded. Start: ~runBWV1069.();  Stop: ~stopBWV1069.() or Cmd+Period.".postln;
)
