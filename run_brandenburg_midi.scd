// Brandenburg MIDI – one file. Open this, run it (Cmd+Enter), it boots and plays.
// Stop: ~stopBrandenburgMIDI.() or Cmd+Period.
s.options.numInputBusChannels = 0;  // avoid input/output sample-rate mismatch
s.waitForBoot {
  // Ensure root group exists (avoids "Group 1 not found")
  s.sendMsg("/g_new", 1, 0, 0);
  s.sync;

  // ---- Paths and part list ----
  ~midiDir = "/Users/kevinwoods/Downloads/brand5-3-mids";
  ~partFiles = ["flute-score.mid", "cello-score.mid", "harpsichord-score.mid"];

  // ---- Native synths ----
  SynthDef(\brand_flute, { |freq = 440, amp = 0.2, sustain = 0.3, gate = 1|
    var env, snd;
    env = EnvGen.kr(Env.linen(0.01, sustain, 0.05, amp, -4), gate, doneAction: 2);
    snd = SinOsc.ar(freq) + (Pulse.ar(freq * 2, 0.2) * 0.2);
    snd = BPF.ar(snd, freq * 2, 0.4);
    Out.ar(0, snd * env ! 2);
  }).add;
  SynthDef(\brand_violin, { |freq = 440, amp = 0.16, sustain = 0.4, gate = 1|
    var env, snd;
    env = EnvGen.kr(Env.linen(0.03, sustain, 0.08, amp, -4), gate, doneAction: 2);
    snd = Saw.ar(freq) + Saw.ar(freq * 2) * 0.25;
    snd = LPF.ar(snd, 2500);
    Out.ar(0, snd * env ! 2);
  }).add;
  SynthDef(\brand_cello, { |freq = 440, amp = 0.2, sustain = 0.5, gate = 1|
    var env, snd;
    env = EnvGen.kr(Env.linen(0.04, sustain, 0.1, amp, -4), gate, doneAction: 2);
    snd = Saw.ar(freq) + Saw.ar(freq * 0.5) * 0.3;
    snd = LPF.ar(snd, 1800);
    Out.ar(0, snd * env ! 2);
  }).add;
  SynthDef(\brand_harpsichord, { |freq = 440, amp = 0.18, sustain = 0.25, gate = 1|
    var env, snd;
    env = EnvGen.kr(Env.perc(0.002, sustain, amp, -6), gate, doneAction: 2);
    snd = Pulse.ar(freq, 0.5) + Pulse.ar(freq * 2, 0.3) * 0.5;
    snd = LPF.ar(snd, 4000);
    Out.ar(0, snd * env ! 2);
  }).add;
  s.sync;

  // ---- MIDI parser (minimal) ----
  ~readMIDIBytes = { |path|
    var f, str, bytes, n;
    f = File(path, "rb");
    if (f.isNil) { ^nil };
    str = f.readAllString;
    f.close;
    if (str.isNil) { ^nil };
    n = str.size;
    bytes = Array.newClear(n);
    (0..(n-1)).do { |i| bytes[i] = str.at(i).ascii & 255 };
    bytes;
  };
  ~readMIDIBytesDebug = { |path| ~readMIDIBytes.(path) };
  ~readVLQ = { |bytes, start, maxPos|
    var val = 0, n = 0, b, limit = maxPos ?? { bytes.size };
    if (start >= limit) { ~readVLQResult = [0, 0] } {
      b = bytes.at(start) & 255; val = b & 127; n = 1;
      if (b & 128 == 0) { ~readVLQResult = [val, n] } {
        if (start + 1 >= limit) { ~readVLQResult = [val, n] } {
          b = bytes[start + 1] & 255; val = (val << 7) + (b & 127); n = 2;
          if (b & 128 == 0) { ~readVLQResult = [val, n] } {
            if (start + 2 >= limit) { ~readVLQResult = [val, n] } {
              b = bytes[start + 2] & 255; val = (val << 7) + (b & 127); n = 3;
              if (b & 128 == 0) { ~readVLQResult = [val, n] } {
                if (start + 3 >= limit) { ~readVLQResult = [val, n] } {
                  b = bytes[start + 3] & 255; val = (val << 7) + (b & 127); n = 4;
                  ~readVLQResult = [val, n];
                };
              };
            };
          };
        };
      };
    };
  };
  ~parseMIDITrack = { |bytes, trackStart, trackLen, division|
    var pos = trackStart, end = trackStart + trackLen, notes = [], lastStatus = 0, tick = 0;
    var getVLQ = { ~readVLQ.(bytes, pos, end); pos = pos + ~readVLQResult[1]; ~readVLQResult[0] };
    var getByte = { var b = if (pos < end) { bytes[pos] & 255 } { 0 }; pos = pos + 1; b };
    var noteOnAt = IdentityDictionary.new;
    var delta, status, note, vel, type, len, prevPos = -1, safeCount = 0, runningDataByte = nil;
    if (trackLen >= 200) {
    while { pos < end } {
      if (pos == prevPos) { pos = pos + 1 };
      prevPos = pos;
    safeCount = safeCount + 1;
    if (safeCount > 20000) { pos = end };
      delta = getVLQ.();
      tick = tick + delta;
      status = getByte.();
      if (status < 128) { runningDataByte = status; status = lastStatus } { runningDataByte = nil };
      if (status >= 128) { lastStatus = status };
      case
      { status >= 144 and: { status <= 159 } } {
        if (runningDataByte.notNil) { note = runningDataByte; vel = getByte.() } { note = getByte.(); vel = getByte.() };
        if (vel > 0) { noteOnAt[note] = tick } { if (noteOnAt[note].notNil) { notes = notes.add([noteOnAt[note], note, tick - noteOnAt[note], 64]); noteOnAt[note] = nil } };
      }
      { status >= 128 and: { status <= 143 } } {
        if (runningDataByte.notNil) { note = runningDataByte; vel = getByte.() } { note = getByte.(); vel = getByte.() };
        if (noteOnAt[note].notNil) { notes = notes.add([noteOnAt[note], note, tick - noteOnAt[note], 64]); noteOnAt[note] = nil };
      }
      { status == 255 } { type = runningDataByte ?? { getByte.() }; len = getByte.(); pos = pos + len }
      { status >= 176 and: { status <= 191 } } { if (runningDataByte.isNil) { getByte.() }; getByte.() }
      { status >= 192 and: { status <= 223 } } { if (runningDataByte.isNil) { getByte.() } }
      { status >= 224 and: { status <= 239 } } { if (runningDataByte.isNil) { getByte.() }; getByte.() }
      { status >= 160 and: { status <= 175 } } { if (runningDataByte.isNil) { getByte.() }; getByte.() }
      { status == 240 or: { status == 247 } } { ~readVLQ.(bytes, pos, end); type = ~readVLQResult[1]; len = ~readVLQResult[0]; pos = pos + type + len }
      { status == 241 or: { status == 243 } } { if (runningDataByte.isNil) { getByte.() } }
      { status == 242 } { if (runningDataByte.isNil) { getByte.() }; getByte.() }
      { true } { };
    };
    };
    noteOnAt.keysDo { |note| if (noteOnAt[note].notNil) { notes = notes.add([noteOnAt[note], note, 384, 64]) } };
    ~lastParsedTrackResult = [notes, division];
  };
  ~parseMIDIFile = { |path|
    var bytes, numTracks, division, tracksNotes, pos, trkLen, res, done = false;
    bytes = ~readMIDIBytesDebug.(path);
    if (bytes.isNil or: { bytes.size < 14 }) { ~lastParsedFileResult = nil } {
      if (bytes[0] != 77 or: { bytes[1] != 84 or: { bytes[2] != 104 or: { bytes[3] != 100 } } }) { ~lastParsedFileResult = nil } {
        numTracks = (bytes[10] << 8) + bytes[11];
        division = (bytes[12] << 8) + bytes[13];
        if (division & 32768 != 0) { ~lastParsedFileResult = nil } {
          tracksNotes = [];
          pos = 14;
          numTracks.do {
            if (done.not) {
              if ((pos + 8) > bytes.size) { ~lastParsedFileResult = [tracksNotes, division]; done = true } {
                if (bytes[pos] != 77 or: { bytes[pos+1] != 84 or: { bytes[pos+2] != 114 or: { bytes[pos+3] != 107 } } }) {
                  pos = pos + 1;
                } {
        trkLen = (bytes[pos+4] << 24) + (bytes[pos+5] << 16) + (bytes[pos+6] << 8) + bytes[pos+7];
        pos = pos + 8;
                  if ((pos + trkLen) <= bytes.size) {
          ~lastParsedTrackResult = nil;
          ~parseMIDITrack.(bytes, pos, trkLen, division);
          res = ~lastParsedTrackResult;
          if (res.notNil and: { res[0].isKindOf(Array) }) { tracksNotes = tracksNotes.add(res[0]) };
                  };
                  pos = pos + trkLen;
                };
              };
            };
          };
          if (done.not) { ~lastParsedFileResult = [tracksNotes, division] };
        };
      };
    };
  };
  ~tickToSec = { |tick, division, usec = 500000| tick * (usec / 1000000.0) / division };
  ~runBrandenburgMIDI = {
    var usec = 500000, raw = [], events = [], synthName, division, minTick;
    ~stopBrandenburgMIDI.();
    s.sendMsg("/g_new", 1, 0, 0);
    s.sync;
    ~brandClock = TempoClock(1, 0);
    Synth(\brand_flute, [\freq, 523.25, \sustain, 0.4, \amp, 0.5], 1);
    "Brandenburg: test beep (C5). If no sound, set Server → Output Device.".postln;
    (~brandenburgParsed ? []).do { |pair|
      var name = pair[0], pParsed = pair[1];
      if (pParsed.notNil and: { pParsed[0].isKindOf(Array) and: { pParsed[0].size > 0 } }) {
        division = pParsed[1];
        synthName = \brand_violin;
        if (name.contains("flute")) { synthName = \brand_flute };
        if (name.contains("cello")) { synthName = \brand_cello };
        if (name.contains("harpsichord")) { synthName = \brand_harpsichord };
        pParsed[0].do { |notes|
          if (notes.isKindOf(Array) and: { notes.notEmpty }) {
            notes.do { |ev| raw = raw.add([ev[0], ev[1], ev[2], ev[3], synthName, division]) };
          };
        };
      };
    };
    if (raw.isEmpty) { "Brandenburg: 0 notes.".postln; ^nil };
    minTick = raw.collect(_[0]).minItem;
    events = raw.collect { |ev|
      [
        ~tickToSec.(ev[0] - minTick, ev[5], usec),
        ev[1],
        ~tickToSec.(ev[2].max(1), ev[5], usec).clip(0.02, 4),
        0.25 * (ev[3] / 127.0),
        ev[4]
      ];
    }.sort({ |a, b| a[0] < b[0] });
    "Brandenburg: % notes. Playing...".format(events.size).postln;
    ~brandPlayRoutine = Routine({
      events.do { |ev|
        (ev[0] - ~brandClock.beats).max(0).wait;
        Synth(ev[4], [\freq, ev[1].midicps, \sustain, ev[2], \amp, ev[3]], 1);
      };
      2.wait;
      Synth(\brand_flute, [\freq, 659.25, \sustain, 0.35, \amp, 0.5], 1);
    }).play(~brandClock);
  };
  ~stopBrandenburgMIDI = { ~brandPlayRoutine !? { ~brandPlayRoutine.stop }; ~brandClock !? { ~brandClock.stop }; s.freeAll };

  // ---- Parse, then start playback (fork so s.sync runs inside a Routine) ----
  s.sync;
  "Brandenburg: parsing MIDI...".postln;
  ~brandenburgParsed = [];
  ~partFiles.do { |name|
    var p = ~midiDir +/+ name;
    if (PathName(p).isFile) {
      ~parseMIDIFile.(p);
      if (~lastParsedFileResult.notNil) {
        ~brandenburgParsed = ~brandenburgParsed.add([name, ~lastParsedFileResult]);
        "  %: % tracks".format(name, ~lastParsedFileResult[0].size).postln;
      };
    };
  };
  "Brandenburg: starting in 1 sec...".postln;
  fork { 1.wait; ~runBrandenburgMIDI.(); };
};
