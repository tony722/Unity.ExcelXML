using System;
using System.Collections.Generic;
using System.Linq;
using Crestron.SimplSharp.CrestronXmlLinq;

namespace AET.Unity.ExcelXml {
  public static class LinqExtensions {
    public static void Each<T>(this IEnumerable<T> ie, Action<T, int> action) {
      var i = 0;
      foreach (var e in ie) action(e, i++);
    }

    public static bool HasAttribute(this XElement element, XName attributeName) {
      if (element.HasAttributes == false) return false;
      return element.Attributes(attributeName).Any();
    }

    public static bool HasElement(this XElement element, XName elementName) {
      if (element.HasElements == false) return false;
      return element.Elements(elementName).Any();
    }

    public static T Previous<T>(this List<T> list, T element) {
      var b = list.ToList();
      b.Reverse();
      return b.SkipWhile(x => !x.Equals(element)).Skip(1).FirstOrDefault();      
    }
  }
}