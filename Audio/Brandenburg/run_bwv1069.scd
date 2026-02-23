s.waitForBoot {
  ("/Users/kevinwoods/Desktop/ActuarialExams/BWV1069_native.scd").load;
  s.sync;
  fork {
    0.5.wait;  // let server tree settle
    ~runBWV1069.();
    "BWV 1069 playing. Stop: Cmd+Period or close sclang.".postln;
  };
};
