# Additional Notes

## Java, React Native, Gradle versions

This demo was tested and runs with React Native 0.62.2, Java 11, and Gradle
3.5.2. Running `make native` will invoke `native.sh`, which uses a fixed version
of React Native 0.62.2 to build and run the demo.

Make sure you have the correct version of Java (11) installed, since 0.62.2 might
not work with newer versions of Java.

## Common Issues

```
ERROR: JAVA_HOME is not set and no 'java' command could be found in your PATH.
```

Add `export JAVA_HOME=<directory>`, replacing `<directory>` with the location of
your Java install, to your `.bashrc` or any other shell that you are using.



