# unisens2excel

[![](https://jitpack.io/v/unisens/unisens2excel.svg)](https://jitpack.io/#unisens/unisens2excel)

unisens2excel is a library to convert a unisens dataset to Excel. All unisens entries with the same samplerate will be included into the excel sheet. Entry descriptions are summarized in a second work sheet.


Add the JitPack repository and the dependency to your build file:

  ```gradle
  repositories {
      maven { url "https://jitpack.io" }
  }
  dependencies {
      compile 'com.github.unisens:unisens2excel:1.0.2'
  }
  ```

Unisens is a **universal data format for multi sensor data**. 
For more information please check the [Unisens website](http://www.unisens.org).

unisens2excel is licenced under the <acronym title="GNU Lesser General Public Licence">LGPL</acronym>.