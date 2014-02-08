openxls - Feb 2014
==================

This is a copy of the LGPL shared openxls library.  It's no longer supported by the owners, but we're using it and are addressing problems as
they crop up.  If you find an issue, please open an issue on Github and/or provide a pull request with a patch.  For any issues logged, please
provide a small simple test case.  Small simple Excel files are also useful in dealing with the multitude of file formats and versions out
there.  Complete stack traces and logs at trace level also help.

OpenXLS supports col/row/cell inserts (and updates formula references when doing so), so this makes it invaluable to us.  This was one of
the reasons we switched from Apache POI (which is also a good library and we have used successfully on several projects).  However, we
invariably end up wrapping POI or OpenXLS since the interfaces don't quite work for us in common situations.  In fact we have one project
that leverages both libraries - something we would like to move away from.  Unfortunately, the functionality provided by OpenXLS and POI
don't fully overlap, so we have to use both libraries behind an SPI facade in order to achieve our goals.

## Plan
Initial commits are code cleanup - bringing the code base up to JDK7 level. We've noticed several bugs while
doing this and will attempt to fix the ones we encounter as we go. The aim is to add unit tests as we fix these issues, but please bear
with us as the code base is around 10 years old and has lots of fixes and workarounds for oddities encountered in the field.

## 2014-02-07
Check Workbook recalc mode, and if set to automatic recalc formulas on Workbook open (similar to how Excel does)

## 2014-01-30
Squashed another bug relating to Mulblank handling

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

