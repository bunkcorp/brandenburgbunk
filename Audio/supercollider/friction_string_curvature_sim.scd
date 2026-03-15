// Friction string / curvature-style ambient instrument study (improved).
// Device strategy: try AirPods first, then fall back to MacBook speakers.

(
~configureFrictionServer = { |outDevice|
    s.options.numInputBusChannels = 0;
    s.options.outDevice = outDevice;
    s.options.inDevice = "MacBook Pro Microphone";
    s.options.blockSize = 512;
    ("friction_string_curvature_sim: trying output -> " ++ outDevice).postln;
};

~startFrictionPatch = {
    SynthDef(\frictionString, {
        |out = 0, freq = 140, pressure = 0.85, bowRate = 0.25, amp = 1.2|

        var bowNoise;
        var wheel;
        var excite;
        var string;
        var body;
        var drift;
        var drift2;
        var feedback;
        var reverb;

        // Mechanical wobble (wheel rotation imperfection).
        drift = LFNoise2.kr(0.15).range(0.97, 1.03);
        drift2 = SinOsc.kr(0.04).range(0.98, 1.02);

        // Rotating wheel contact profile.
        wheel = LFTri.kr(bowRate).range(0, 1);

        // Stick-slip style friction approximation.
        bowNoise = PinkNoise.ar * pressure;
        bowNoise = bowNoise * (wheel + 0.1);
        bowNoise = LeakDC.ar(bowNoise);

        // Nonlinear bow interaction.
        excite = tanh(bowNoise * 6);
        // Continuous low-energy injection keeps long drones alive.
        excite = excite + (SinOsc.ar(freq * 0.5) * 0.02);

        // String resonance.
        string = CombL.ar(
            excite,
            0.2,
            (1 / (freq * drift * drift2)).clip(0.001, 0.2),
            10
        );

        // Multi-mode resonant body (wood plate simulation).
        body =
            BPF.ar(string, freq * 1.0, 0.2) +
            BPF.ar(string, freq * 1.61, 0.2) +
            BPF.ar(string, freq * 2.32, 0.25) +
            BPF.ar(string, freq * 3.73, 0.3);
        // Add a little dry string back so the tone stays present.
        body = body + (LPF.ar(string, 3500) * 0.35);

        // Subtle sustain feedback.
        feedback = LocalIn.ar(1);
        body = body + (feedback * 0.22);
        LocalOut.ar(LeakDC.ar(Mix(body)));

        // Ambient diffusion (plugin-safe fallback to FreeVerb2).
        reverb = FreeVerb2.ar(body, body, 0.6, 0.92, 0.4);

        Out.ar(out, Limiter.ar((body + reverb) * amp ! 2, 0.95));
    }).add;

    SynthDef(\granularCloud, {
        |out = 0, inChan = 0, trigRate = 6, grainDur = 0.25, amp = 0.5|

        var source;
        var grains;
        var spread;
        var reverb;

        source = SoundIn.ar(inChan);

        // Stereo motion.
        spread = LFNoise1.kr(0.2);

        grains = GrainIn.ar(
            2,
            Impulse.kr(trigRate),
            grainDur,
            source,
            spread
        );

        reverb = FreeVerb2.ar(grains[0], grains[1], 0.75, 0.92, 0.45);

        Out.ar(out, reverb * amp);
    }).add;

    s.sync;

    "friction_string_curvature_sim: server booted".postln;

    // Short test tone.
    {
        var env = Env.perc(0.001, 0.18).ar(doneAction: 2);
        var sig = SinOsc.ar(660) * 0.15 * env;
        Out.ar(0, sig ! 2);
    }.play;

    // Play friction string.
    x = Synth(\frictionString);
    "friction_string_curvature_sim: friction synth started".postln;

    // Optional controls:
    // x.set(\freq, 70);
    // x.set(\pressure, 0.9);
    // x.set(\bowRate, 0.1);

    // Optional granular layer:
    // y = Synth(\granularCloud);

    // Stop:
    // x.free;
    // y.free;
};

~bootFrictionWithFallback = { |primaryOut, fallbackOut|
    ~configureFrictionServer.(primaryOut);
    s.waitForBoot(
        {
            ~startFrictionPatch.();
        },
        100,
        {
            ("friction_string_curvature_sim: failed on -> " ++ primaryOut).warn;
            if(fallbackOut.notNil) {
                AppClock.sched(0.5, {
                    ~bootFrictionWithFallback.(fallbackOut, nil);
                    nil
                });
            };
        }
    );
};

~bootFrictionWithFallback.("Kevin’s AirPods P", "MacBook Pro Speakers");
)
