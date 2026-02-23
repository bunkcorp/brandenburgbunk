/*
  Brandenburg DnB – SP1200 French Touch Engine
  Main: load modules, busses, DSP/FX chain, OSC control, demo.
  Run after server is booted: s.waitForBoot { (path +/+ "main.scd").load };
*/
(
var projectDir = PathName(thisProcess.nowExecutingPath).pathOnly;

"Loading Brandenburg DnB engine...".postln;

// State in environment so OSCdefs and ~startDSP stay valid after block
~brandenburgSourceBus = Bus.audio(s, 2);
~brandenburgDspBus = Bus.audio(s, 2);
~brandenburgFxBus = Bus.audio(s, 2);
~brandenburgFxBus2 = Bus.audio(s, 2);
~brandenburgKickTrigBus = Bus.control(s, 1);
~brandenburgMainOut = 0;
~brandenburgDspMode = 0;
~brandenburgDnbRoutine = nil;
~brandenburgSp1200Synth = nil;
~brandenburgS760Synth = nil;
~brandenburgReverbSynth = nil;
~brandenburgDelaySynth = nil;
~brandenburgChorusSynth = nil;
~brandenburgSidechainSynth = nil;

// Load submodules
(projectDir +/+ "engine/SampleLoader.scd").load;
(projectDir +/+ "engine/SliceSampler.scd").load;
(projectDir +/+ "dsp/SP1200.scd").load;
(projectDir +/+ "dsp/S760.scd").load;
(projectDir +/+ "fx/ReverbDelayChain.scd").load;
(projectDir +/+ "patterns/DnBEngine.scd").load;

s.sync;

~brandenburgDspGroup = Group.tail(s);
~brandenburgFxGroup = Group.after(~brandenburgDspGroup);

~startDSP = {
  ~brandenburgSp1200Synth.free;
  ~brandenburgS760Synth.free;
  if (~brandenburgDspMode == 1) {
    ~brandenburgSp1200Synth = Synth(\sp1200_stereo, [\inBus, ~brandenburgSourceBus, \outBus, ~brandenburgDspBus], ~brandenburgDspGroup);
  } {
    ~brandenburgS760Synth = Synth(\s760_stereo, [\inBus, ~brandenburgSourceBus, \outBus, ~brandenburgDspBus], ~brandenburgDspGroup);
  };
};

~startFX = {
  ~brandenburgReverbSynth = Synth(\fx_ethereal_reverb, [\inBus, ~brandenburgDspBus, \outBus, ~brandenburgFxBus, \decay, 5.5, \wet, 0.28], ~brandenburgFxGroup, \addToHead);
  ~brandenburgDelaySynth = Synth(\fx_pingpong_delay, [\inBus, ~brandenburgFxBus, \outBus, ~brandenburgFxBus2, \time, 0.272, \feedback, 0.4, \wet, 0.2], ~brandenburgFxGroup, \addToTail);
  ~brandenburgChorusSynth = Synth(\fx_chorus, [\inBus, ~brandenburgFxBus2, \outBus, ~brandenburgFxBus2, \rate, 0.3, \depth, 0.004], ~brandenburgFxGroup, \addToTail);
  ~brandenburgSidechainSynth = Synth(\fx_sidechain, [\inBus, ~brandenburgFxBus2, \outBus, ~brandenburgMainOut, \triggerBus, ~brandenburgKickTrigBus, \release, 0.25], ~brandenburgFxGroup, \addToTail);
};

~playSlice = { |sliceIndex = 0, speed = 1, reverse = 0, duration = 0.5|
  var buf = ~brandenburgBuf;
  var frames = ~getSliceFrames.(sliceIndex);
  if (buf.notNil) {
    Synth(\brandenburg_slice, [
      \bufnum, buf, \startFrame, frames[0], \numFrames, frames[1],
      \speed, speed, \reverse, reverse, \outBus, ~brandenburgSourceBus, \amp, 0.7
    ], ~brandenburgDspGroup, \addToHead);
  } { "Load Brandenburg sample first: ~loadBrandenburg(path)".postln };
};

~bpm = 172;
~stepSec = 60 / ~bpm / 4;
~runDnBDemo = {
  ~brandenburgDnbRoutine !? { ~brandenburgDnbRoutine.stop };
  ~brandenburgDnbRoutine = Routine({
    loop {
      8.do { |i|
        if ([0, 4].includes(i)) {
          Synth(\dnb_kick, [\outBus, ~brandenburgMainOut], ~brandenburgDspGroup, \addToHead);
          ~brandenburgKickTrigBus.set(1);
          { ~brandenburgKickTrigBus.set(0) }.defer(0.05);
        };
        if (i == 4) { Synth(\dnb_snare, [\outBus, ~brandenburgMainOut], ~brandenburgDspGroup, \addToHead) };
        if ([2, 6].includes(i)) { Synth(\dnb_ghost_snare, [\outBus, ~brandenburgMainOut], ~brandenburgDspGroup, \addToHead) };
        4.do { Synth(\dnb_hat, [\outBus, ~brandenburgMainOut], ~brandenburgDspGroup, \addToHead) };
        ~stepSec.wait;
      };
    };
  }).play(~dnbClock);
};
~stopDnB = { ~brandenburgDnbRoutine !? { ~brandenburgDnbRoutine.stop } };

OSCdef(\brandenburg_dsp_mode, { |msg| ~brandenburgDspMode = msg[1].asInteger.clip(0, 1); ~startDSP.(); }, \brandenburg/dsp_mode);
OSCdef(\brandenburg_rev_decay, { |msg| ~brandenburgReverbSynth !? { ~brandenburgReverbSynth.set(\decay, msg[1].asFloat) } }, \brandenburg/fx/rev_decay);
OSCdef(\brandenburg_rev_wet, { |msg| ~brandenburgReverbSynth !? { ~brandenburgReverbSynth.set(\wet, msg[1].asFloat) } }, \brandenburg/fx/rev_wet);
OSCdef(\brandenburg_delay_fb, { |msg| ~brandenburgDelaySynth !? { ~brandenburgDelaySynth.set(\feedback, msg[1].asFloat) } }, \brandenburg/fx/delay_fb);
OSCdef(\brandenburg_play_slice, { |msg| ~playSlice.(msg[1].asInteger, msg[2].asFloat, msg[3].asInteger, msg[4].asFloat) }, \brandenburg/play_slice);
OSCdef(\brandenburg_filter, { |msg| ~brandenburgS760Synth !? { ~brandenburgS760Synth.set(\cutoff, msg[1].asFloat) } }, \brandenburg/filter_cutoff);
OSCdef(\brandenburg_sp_drive, { |msg| ~brandenburgSp1200Synth !? { ~brandenburgSp1200Synth.set(\drive, msg[1].asFloat) } }, \brandenburg/sp_drive);

~startDSP.();
~startFX.();

"Brandenburg DnB ready. Busses: source -> DSP -> reverb -> delay -> chorus -> sidechain -> out.".postln;
"OSC: /brandenburg/dsp_mode (0=S760, 1=SP1200), /brandenburg/play_slice i speed reverse dur, /brandenburg/fx/rev_decay, rev_wet, delay_fb, /brandenburg/filter_cutoff, /brandenburg/sp_drive.".postln;
"Global: ~loadBrandenburg(path), ~playSlice(i,speed,rev,dur), ~runDnBDemo, ~stopDnB.".postln;
)
