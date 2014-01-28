openxls
=======

This is a copy of the LGPL shared openxls library.  It's no longer supported by the owners, but we're using it and are fixing things as we
encounter any problems.  If you find an issue, please open an issue on Github and/or provide a pull request.  For any issues logged, please
provide a small simple test case.  Small simple Excel files are also useful in dealing with the multitude of file formats and versions out
there.

## Plan
Initial commits are going to be code cleanup - bringing the code base up to JDK7 level. We've noticed several bugs while
doing this and will attempt to fix the ones we encounter as we go. The aim is to add unit tests as we fix these issues, but please bear
with us as the code base is many years old and has lots of fixes and workarounds for oddities encountered in the field.  We're also moving
the logging to SLF4J, and ensuring that all exceptions are logged at WARN or ERROR level (depending on severity).

A big issue for us right now is the corruption of one of the two TreeMaps held within Boundsheet - this is an ordered tree, and instance
variables that are part of the ordering are being modified in place.  This prevents the nodes from being found in subsequent searches and
deletions.

## 2014-01-28
Snapshots are now being built on our CI server (TeamCity) and published to Sonatype's OSS Snapshot Repo.  We will package, sign and publish
a release build shortly (which will be auto-synced from Sonatype's Release Repo to Maven Central)

In the meantime, you can access the snapshots by adding the following to your Maven pom.xml:

```xml
    <repositories>

        <repository>
            <id>Sonatype Snapshots</id>
            <name>Sonatype OSS Snapshots</name>
            <url>https://oss.sonatype.org/content/repositories/snapshots/</url>
            <snapshots>
                <enabled>true</enabled>
            </snapshots>
            <releases>
                <enabled>false</enabled>
            </releases>
        </repository>

    </repositories>
```

and reference this dependency:

```xml
        <dependency>
            <groupId>org.openxls</groupId>
            <artifactId>openxls</artifactId>
            <version>12.0-SNAPSHOT</version>
        </dependency>
```

## 2014-01-25
Replaced an internal TreeMap of Boundsheet that handled cells in columns with an alternative unit tested implementation.  This fixes a bug
we were hitting relating to overrunning an array since the cells-by-col method was returning an incorrect set of Cells in some cases.

## 2014-01-22
Logging has been moved to SLF4J.  There will likely be many log lines that are being recorded at the wrong level (i.e. too much noise).
Please feel free to send a pull request with your level changes, or log an issue with the file name and line number, the current level and
what you think it should be.  If you think there should be additional log lines (in currently empty Exception catch blocks) please feel
free to suggest additions.

