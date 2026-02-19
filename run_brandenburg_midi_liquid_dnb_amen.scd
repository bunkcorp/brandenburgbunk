// Brandenburg Liquid DnB — Amen version (exact break rhythm, native synthesis)
// Requires: Quarks.install("wslib") then recompile (Cmd+K). Stop: ~stopBrandenburgMIDI.()
//
// Amen: kick 1&9, snare 2&4 (0.8), ghosts (0.3), 16th hats. Optional: \lag, Pwhite(-0.005,0.005) for micro-timing.

s.options.numInputBusChannels = 0;

s.waitForBoot {

    ~midiDir = "/Users/kevinwoods/Downloads/brand5-3-mids";
    ~partFiles = ["flute-score.mid", "cello-score.mid", "harpsichord-score.mid"];

    ~verbBus   = Bus.audio(s, 2);
    ~masterBus = Bus.audio(s, 2);
    ~drumBus   = Bus.audio(s, 2);
    ~triggerBus= Bus.audio(s, 1);
    ~levelBus  = Bus.control(s, 1);

    //////////////////////////////
    // REVERB + BUILD SWEEP
    //////////////////////////////

    SynthDef(\lushVerb, { |verbBus=0, masterBus=0, intro=2, build=8, hold=90|
        var dry, chorus, sweep, wet;
        dry = In.ar(verbBus, 2);
        chorus = dry + DelayC.ar(dry, 0.03, LFTri.kr(0.2).range(0.005, 0.018)) * 0.3;
        sweep = EnvGen.kr(Env([400,400,6000,6000], [intro,build,hold], \lin));
        chorus = LPF.ar(chorus, sweep);
        wet = FreeVerb2.ar(chorus[0], chorus[1], 0.5, 0.9, 0.3);
        Out.ar(masterBus, dry + wet * 0.6);
    }).add;

    SynthDef(\sidechainTrigger, { |triggerBus=0, tempo=2|
        var trig = Impulse.kr(tempo);
        var env  = EnvGen.ar(Env.perc(0.001, 0.18), trig);
        Out.ar(triggerBus, env);
    }).add;

    SynthDef(\multibandSidechain, {
        |masterBus=0, drumBus=0, triggerBus=0, levelBus=0, duck=0.5|
        var sig, drums, sc, level, low, mid, gain;
        sig   = In.ar(masterBus, 2);
        drums = In.ar(drumBus, 2);
        sc    = In.ar(triggerBus, 1);
        level = In.kr(levelBus, 1);
        low   = LPF.ar(sig, 120);
        mid   = HPF.ar(sig, 120);
        gain  = 1 - (sc * duck);
        mid   = mid * gain;
        Out.ar(0, ((low + mid) + drums) * level);
    }).add;

    SynthDef(\arrangementLevel, { |levelBus=0, intro=2, build=8, hold=90, breakdown=10|
        var env = Env([0,0,1,1,0], [intro,build,hold,breakdown], \sin);
        Out.kr(levelBus, EnvGen.kr(env, doneAction:2));
    }).add;

    //////////////////////////////
    // ORCHESTRA
    //////////////////////////////

    // Self-freeing (sustain-based) so no /n_set release → no "Node not found"
    SynthDef(\brand_flute, { |freq=440, amp=0.3, sustain=0.3, verbBus=0|
        var env, sig, width;
        env = EnvGen.kr(Env([0,1,1,0], [0.05, 0.5, 0.2]), 1, sustain.max(0.02)/0.75, doneAction:2);
        sig = Saw.ar(freq * [1,1.003]) * 0.5;
        sig = LPF.ar(sig.sum, 6000);
        width = SinOsc.kr(0.05).range(0.3,0.8);
        Out.ar(verbBus, Splay.ar(sig * env * amp, width));
    }).add;

    SynthDef(\brand_cello, { |freq=220, amp=0.4, sustain=0.5, verbBus=0|
        var env, sig, width;
        env = EnvGen.kr(Env([0,1,1,0], [0.1, 0.5, 0.5]), 1, sustain.max(0.02)/1.1, doneAction:2);
        sig = Saw.ar(freq * [1,1.01]).sum;
        sig = LPF.ar(sig, 1200);
        width = SinOsc.kr(0.05).range(0.3,0.8);
        Out.ar(verbBus, Splay.ar(sig * env * amp, width));
    }).add;

    SynthDef(\brand_harpsichord, { |freq=440, amp=0.25, sustain=0.3, verbBus=0|
        var env, sig, width;
        env = EnvGen.kr(Env.perc(0.003, 1), 1, sustain.max(0.02), doneAction:2);
        sig = HPF.ar(Pulse.ar(freq,0.5), 500);
        width = SinOsc.kr(0.05).range(0.3,0.8);
        Out.ar(verbBus, Splay.ar(sig * env * amp, width));
    }).add;

    SynthDef(\reeseBass, { |freq=55, amp=0.7, gate=1, masterBus=0|
        var env, detune, reese, sub, sig;
        env = EnvGen.kr(Env.adsr(0.01,0.3,0.8,0.2), gate, doneAction:2);
        detune = [0.99,1.01];
        reese  = Saw.ar(freq * detune).sum;
        reese = LPF.ar(reese, 800);
        reese = reese.tanh;
        sub = SinOsc.ar(freq, 0, 0.5);
        sig = LeakDC.ar(reese + sub);
        Out.ar(masterBus, Pan2.ar(sig * env * amp));
    }).add;

    //////////////////////////////
    // DRUMS — Exact Amen break (kick 1&9, snare 2&4, ghosts, 16th hats)
    //////////////////////////////

    SynthDef(\amenKick, { |drumBus=0|
        var env = EnvGen.ar(Env.perc(0.001, 0.18), doneAction:2);
        var body = SinOsc.ar(55, 0, env);
        var click = WhiteNoise.ar * EnvGen.ar(Env.perc(0.001, 0.02));
        Out.ar(drumBus, Pan2.ar((body + (click * 0.2)) * 0.9));
    }).add;

    SynthDef(\amenSnare, { |drumBus=0, amp=0.6|
        var env = EnvGen.ar(Env.perc(0.005, 0.2), doneAction:2);
        var noise = WhiteNoise.ar * env;
        var tone = SinOsc.ar(180, 0, env * 0.3);
        Out.ar(drumBus, Pan2.ar((noise + tone) * amp));
    }).add;

    SynthDef(\amenHat, { |drumBus=0, amp=0.3|
        var env = EnvGen.ar(Env.perc(0.001, 0.05), doneAction:2);
        var sig = HPF.ar(WhiteNoise.ar, 6000);
        Out.ar(drumBus, Pan2.ar(sig * env * amp));
    }).add;

    s.sync;

    TempoClock.default.tempo = 172/60;
    // Swing: SC TempoClock has no built-in swing; use Pbind \dur with Pseq of slightly staggered values if desired

    ~levelBus.set(0);

    ~verbSynth = Synth(\lushVerb, [
        \verbBus, ~verbBus.index,
        \masterBus, ~masterBus.index,
        \intro, 2, \build, 8, \hold, 90
    ], s, \addToTail);

    ~triggerSynth = Synth(\sidechainTrigger, [
        \triggerBus, ~triggerBus.index,
        \tempo, TempoClock.default.tempo
    ], s, \addToTail);

    ~sidechain = Synth(\multibandSidechain, [
        \masterBus, ~masterBus.index,
        \drumBus, ~drumBus.index,
        \triggerBus, ~triggerBus.index,
        \levelBus, ~levelBus.index,
        \duck, 0.4   // liquid: try 0.45–0.55 for more pump
    ], s, \addToTail);

    //////////////////////////////
    // MIDI LOAD
    //////////////////////////////

    ~patterns = [];
    ~partFiles.do { |name|
        var path = ~midiDir +/+ name;
        if(PathName(path).isFile){
            var midi = SimpleMIDIFile.read(path);
            ~patterns = ~patterns.add([name, midi.p]);
            ("Loaded: " ++ name).postln;
        };
    };

    ~makePart = { |name, pat|
        var synth = \brand_flute;
        if(name.contains("cello")) { synth = \brand_cello };
        if(name.contains("harpsichord")) { synth = \brand_harpsichord };
        Pbindf(pat,
            \instrument, synth,
            \dur, Pkey(\dur) * 0.25,
            \sustain, Pkey(\dur) * 0.95,
            \amp, (Pkey(\amp) ? ((Pkey(\vel) ? 100) / 127)) * 0.5,
            \verbBus, ~verbBus.index,
            \hasGate, false
        );
    };

    ~orchestra = Ppar(~patterns.collect{|p| ~makePart.(p[0],p[1])});

    ~bass = Pbind(
        \instrument, \reeseBass,
        \scale, Scale.minor,
        \degree, Pseq([0,-2,-3,-5], inf),
        \octave, 2,
        \dur, 0.5,
        \masterBus, ~masterBus.index
    );

    //////////////////////////////
    // AMEN PATTERN — Exact 1-bar rhythm (steps 1–16)
    // Kick: 1 & 9   Snare: 5 & 13 (backbeat)   Ghost: 2,4,7,10,12,15
    //////////////////////////////

    ~kick = Pbind(
        \instrument, \amenKick,
        \dur, 0.25,
        \isRest, Pseq([0, 1, 1, 1,  1, 1, 1, 1,  0, 1, 1, 1,  1, 1, 1, 1], inf),
        \midinote, Pfunc { |ev| if(ev[\isRest] != 0) { \rest } { 60 } },
        \drumBus, ~drumBus.index
    );

    ~snare = Pbind(
        \instrument, \amenSnare,
        \dur, 0.25,
        \amp, Pseq([0, 0.3, 0, 0,  0.8, 0, 0.3, 0,  0, 0.3, 0, 0,  0.8, 0, 0.3, 0], inf),
        \isRest, Pfunc { |ev| ev[\amp] <= 0 },
        \midinote, Pfunc { |ev| if(ev[\isRest]) { \rest } { 60 } },
        \drumBus, ~drumBus.index
    );

    ~hats = Pbind(
        \instrument, \amenHat,
        \dur, 0.25,
        \amp, Pwhite(0.2, 0.35),
        \drumBus, ~drumBus.index
    );

    ~drums = Ppar([~kick, ~snare, ~hats]);

    ~score = Ppar([~orchestra, ~bass, ~drums]);

    ~startBrandenburg = {
        ~player = ~score.play;
        Synth(\arrangementLevel, [\levelBus, ~levelBus.index], s, \addToTail);
        "Brandenburg Liquid DnB Amen @ 172 BPM (syncopated kick, ghost snares)".postln;
        nil
    };

    ~stopBrandenburgMIDI = {
        ~player.stop;
        fork { 2.wait; s.freeAll; "Stopped.".postln };
        nil
    };

    ~startBrandenburg.();
    nil
};
~stopBrandenburgMIDI.();
