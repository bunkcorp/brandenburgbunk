/*
  Brandenburg from MIDI – play brand5-3-mids .mid files with native SuperCollider synths.
  No downloads: parses MIDI in SC, maps parts to synths (flute, violin, cello, etc.).
  Run after s.boot. Start: ~runBrandenburgMIDI.();  Stop: ~stopBrandenburgMIDI.()
*/
(
~midiDir = "/Users/kevinwoods/Downloads/brand5-3-mids";
// Quick load (3 parts, ~2–3 sec). For full ensemble, use: ~partFiles = ["flute-score.mid", "violin-score.mid", "violinP-score.mid", "viola-score.mid", "cello-score.mid", "violone-score.mid", "harpsichord-score.mid"];
~partFiles = ["flute-score.mid", "cello-score.mid", "harpsichord-score.mid"];
~scoreFile = "score.mid";

// ---- Native synths (no samples) ----
SynthDef(\brand_flute, { |freq = 440, amp = 0.2, sustain = 0.3, gate = 1|
  var env = EnvGen.kr(Env.linen(0.01, sustain, 0.05, amp, -4), gate, doneAction: 2);
  var snd = SinOsc.ar(freq) + (Pulse.ar(freq * 2, 0.2) * 0.2);
  snd = BPF.ar(snd, freq * 2, 0.4);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\brand_oboe, { |freq = 440, amp = 0.18, sustain = 0.35, gate = 1|
  var env = EnvGen.kr(Env.linen(0.02, sustain, 0.05, amp, -4), gate, doneAction: 2);
  var snd = Saw.ar(freq) + Pulse.ar(freq, 0.3) * 0.4;
  snd = BPF.ar(snd, freq * 2, 0.3);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\brand_violin, { |freq = 440, amp = 0.16, sustain = 0.4, gate = 1|
  var env = EnvGen.kr(Env.linen(0.03, sustain, 0.08, amp, -4), gate, doneAction: 2);
  var snd = Saw.ar(freq) + Saw.ar(freq * 2) * 0.25;
  snd = LPF.ar(snd, 2500);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\brand_viola, { |freq = 440, amp = 0.17, sustain = 0.4, gate = 1|
  var env = EnvGen.kr(Env.linen(0.03, sustain, 0.08, amp, -4), gate, doneAction: 2);
  var snd = Saw.ar(freq) + Saw.ar(freq * 2) * 0.28;
  snd = LPF.ar(snd, 2200);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\brand_cello, { |freq = 440, amp = 0.2, sustain = 0.5, gate = 1|
  var env = EnvGen.kr(Env.linen(0.04, sustain, 0.1, amp, -4), gate, doneAction: 2);
  var snd = Saw.ar(freq) + Saw.ar(freq * 0.5) * 0.3;
  snd = LPF.ar(snd, 1800);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\brand_violone, { |freq = 440, amp = 0.22, sustain = 0.5, gate = 1|
  var env = EnvGen.kr(Env.linen(0.04, sustain, 0.12, amp, -4), gate, doneAction: 2);
  var snd = Saw.ar(freq) + SinOsc.ar(freq) * 0.4;
  snd = LPF.ar(snd, 1200);
  Out.ar(0, snd * env ! 2);
}).add;

SynthDef(\brand_harpsichord, { |freq = 440, amp = 0.18, sustain = 0.25, gate = 1|
  var env = EnvGen.kr(Env.perc(0.002, sustain, amp, -6), gate, doneAction: 2);
  var snd = Pulse.ar(freq, 0.5) + Pulse.ar(freq * 2, 0.3) * 0.5;
  snd = LPF.ar(snd, 4000);
  Out.ar(0, snd * env ! 2);
}).add;

s.sync;

// ---- Minimal MIDI parser ----
// Read file bytes. Returns array of 0-255.
~readMIDIBytes = { |path|
  var f, str, bytes, n;
  f = File(path, "rb");
  if (f.isNil) { "File not found: %".format(path).postln; ^nil };
  str = f.readAllString;
  f.close;
  if (str.isNil) { ^nil };
  n = str.size;
  bytes = Array.newClear(n);
  (0..(n-1)).do { |i| bytes[i] = str.at(i).ascii };
  bytes;
};

// Parse VLQ (variable-length quantity), return [value, numBytesConsumed]
~readVLQ = { |bytes, start|
  var val = 0, i = start, b;
  while { i < bytes.size } {
    b = bytes[i];
    val = (val << 7) + (b & 127);
    i = i + 1;
    if (b & 128 == 0) { ^[val, i - start] };
  };
  ^[val, i - start];
};

// Parse one track; return [notes, division]. notes = [[startTick, noteNum, durTicks, vel], ...]
~parseMIDITrack = { |bytes, trackStart, trackLen, division|
  var pos = trackStart;
  var end = trackStart + trackLen;
  var notes = [];
  var lastStatus = 0;
  var tick = 0;
  var usecPerQuarter = 500000;
  var tempoEvents = [];
  var getVLQ = { var r = ~readVLQ.(bytes, pos); pos = pos + r[1]; r[0] };
  var getByte = { var b = bytes[pos]; pos = pos + 1; b };
  var noteOnAt = IdentityDictionary.new;
  var delta, status, note, vel, type, len;
  var prevPos = -1;
  var safeCount = 0;
  while { pos < end } {
    if (pos == prevPos) { pos = pos + 1 };  // force advance if stuck
    prevPos = pos;
    safeCount = safeCount + 1;
    if (safeCount > 500000) { "  Track parse limit (500k events), stopping.".postln; pos = end };
    delta = getVLQ.();
    tick = tick + delta;
    status = getByte.();
    if (status < 128) { pos = pos - 1; status = lastStatus };
    if (status >= 128) { lastStatus = status };
    case
    { status >= 144 and: { status <= 159 } } {
      note = getByte.();
      vel = getByte.();
      if (vel > 0) {
        noteOnAt[note] = tick;
      } {
        if (noteOnAt[note].notNil) {
          notes = notes.add([noteOnAt[note], note, tick - noteOnAt[note], 64]);
          noteOnAt[note] = nil;
        };
      };
    }
    { status >= 128 and: { status <= 143 } } {
      note = getByte.();
      vel = getByte.();
      if (noteOnAt[note].notNil) {
        notes = notes.add([noteOnAt[note], note, tick - noteOnAt[note], 64]);
        noteOnAt[note] = nil;
      };
    }
    { status == 255 } {
      type = getByte.();
      len = getByte.();
      if (type == 81 and: { len == 3 }) {
        usecPerQuarter = (bytes[pos] << 16) + (bytes[pos+1] << 8) + bytes[pos+2];
        tempoEvents = tempoEvents.add([tick, usecPerQuarter]);
      };
      pos = pos + len;
    }
    { status >= 176 and: { status <= 191 } } { pos = pos + 2 }  // controller
    { status >= 192 and: { status <= 207 } } { pos = pos + 1 }  // program
    { status >= 208 and: { status <= 223 } } { pos = pos + 1 }  // channel aftertouch
    { status >= 224 and: { status <= 239 } } { pos = pos + 2 }  // pitch bend
    { status >= 160 and: { status <= 175 } } { pos = pos + 2 }  // poly aftertouch
    { status == 240 or: { status == 247 } } {  // sysex: VLQ length then skip that many bytes
      type = ~readVLQ.(bytes, pos)[1];
      len = ~readVLQ.(bytes, pos)[0];
      pos = pos + type + len;
    }
    { status == 241 or: { status == 243 } } { pos = pos + 1 }  // F1, F3: 1 data byte
    { status == 242 } { pos = pos + 2 }  // F2: 2 data bytes
    { true } { };  // F8–FE real-time: no data (already advanced by getByte)
  };
  // close any open notes at end
  noteOnAt.keysDo { |note|
    if (noteOnAt[note].notNil) {
      notes = notes.add([noteOnAt[note], note, 384, 64]);  // 1 beat default
    };
  };
  ^[notes, division, tempoEvents];
};

// Parse full MIDI file; return [tracksNotes, division]. tracksNotes = [[[startTick, note, dur, vel], ...], ...]
~parseMIDIFile = { |path|
  var bytes, headerLen, format, numTracks, division, tracksNotes, pos, trkLen, res;
  bytes = ~readMIDIBytes.(path);
  if (bytes.isNil or: { bytes.size < 14 }) { ^nil };
  if (bytes[0] != 77 or: { bytes[1] != 84 or: { bytes[2] != 104 or: { bytes[3] != 100 } } }) {
    "Not MThd".postln; ^nil
  };
  headerLen = (bytes[4] << 24) + (bytes[5] << 16) + (bytes[6] << 8) + bytes[7];
  format = (bytes[8] << 8) + bytes[9];
  numTracks = (bytes[10] << 8) + bytes[11];
  division = (bytes[12] << 8) + bytes[13];
  if (division & 32768 != 0) { "SMPTE division not supported".postln; ^nil };
  tracksNotes = [];
  pos = 14;
  numTracks.do {
    if (pos + 8 > bytes.size) { ^[tracksNotes, division] };
    if (bytes[pos] != 77 or: { bytes[pos+1] != 84 or: { bytes[pos+2] != 114 or: { bytes[pos+3] != 107 } } }) {
      pos = pos + 1;
    } {
      trkLen = (bytes[pos+4] << 24) + (bytes[pos+5] << 16) + (bytes[pos+6] << 8) + bytes[pos+7];
      pos = pos + 8;
      res = ~parseMIDITrack.(bytes, pos, trkLen, division);
      if (res[0].isKindOf(Array)) { tracksNotes = tracksNotes.add(res[0]) };
      pos = pos + trkLen;
    };
  };
  ^[tracksNotes, division];
};

// Ticks to seconds: tick * (usecPerQuarter/1e6) / division
~tickToSec = { |tick, division, usecPerQuarter = 500000|
  ^tick * (usecPerQuarter / 1000000.0) / division
};

// ---- Play one track with a given synth ----
// targetGroup: 1 = root we create so server definitely has a node to add to
~playTrack = { |notes, division, synthName, clock, targetGroup = 1, timeOffset = 0|
  var usec = 500000;
  var startTick, note, durTicks, vel, startSec, durSec;
  if (notes.isEmpty) { ^nil };
  Routine({
    notes.sort({ |a, b| a[0] < b[0] }).do { |ev|
      startTick = ev[0];
      note = ev[1];
      durTicks = ev[2].max(1);
      vel = ev[3] / 127.0;
      startSec = ~tickToSec.(startTick, division, usec) - timeOffset;
      if (startSec < 0) { startSec = 0 };
      durSec = ~tickToSec.(durTicks, division, usec).clip(0.02, 4);
      ~brandClock.sched(startSec, {
        Synth(synthName, [\freq, note.midicps, \sustain, durSec, \amp, 0.2 * vel], targetGroup);
        nil;
      });
    };
  }).play(clock);
};

// ---- Main: load score or parts and play ----
~runBrandenburgMIDI = {
  var path, parsed, tracks, div, synths, p, pParsed, synthName, minTick, totalNotes;
  "Brandenburg: starting...".postln;
  ~stopBrandenburgMIDI.();
  s.sendMsg("/g_new", 1, 0, 0);
  s.sync;
  ~brandClock = TempoClock(1, 0);
  ~brandRoutines = [];
  // Use part files only (score.mid is large and can hang the parser)
  "Brandenburg: loading part files...".postln;
  ~partFiles.do { |name|
    p = ~midiDir +/+ name;
    if (PathName(p).isFile) {
      "  Parsing % ...".format(name).postln;
      pParsed = ~parseMIDIFile.(p);
      if (pParsed.notNil and: { pParsed[0].isKindOf(Array) and: { pParsed[0].size > 0 } }) {
        synthName = \brand_violin;
        if (name.contains("flute")) { synthName = \brand_flute };
        if (name.contains("oboe")) { synthName = \brand_oboe };
        if (name.contains("viola")) { synthName = \brand_viola };
        if (name.contains("cello")) { synthName = \brand_cello };
        if (name.contains("violone")) { synthName = \brand_violone };
        if (name.contains("harpsichord")) { synthName = \brand_harpsichord };
        if (name.contains("violinP")) { synthName = \brand_violin };
        pParsed[0].do { |trackNotes, ti|
          if (trackNotes.isKindOf(Array) and: { trackNotes.notEmpty }) {
            minTick = trackNotes.collect(_[0]).minItem;
            ~brandRoutines = ~brandRoutines.copy.add(~playTrack.(trackNotes, pParsed[1], synthName, ~brandClock, 1, ~tickToSec.(minTick, pParsed[1], 500000)));
          };
        };
        "  Loaded % (% notes)".format(name, pParsed[0].sum(_.size)).postln;
      };
    };
  };
  "Brandenburg MIDI playing. Stop: ~stopBrandenburgMIDI.().".postln;
};

~stopBrandenburgMIDI = {
  ~brandRoutines !? { ~brandRoutines.do(_.stop) };
  ~brandClock !? { ~brandClock.stop };
  s.freeAll;
};

"Brandenburg MIDI loaded. Start: ~runBrandenburgMIDI.();  Stop: ~stopBrandenburgMIDI.().".postln;
)
