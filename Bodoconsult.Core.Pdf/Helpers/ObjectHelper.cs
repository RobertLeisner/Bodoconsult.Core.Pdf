// Copyright (c) Bodoconsult EDV-Dienstleistungen GmbH. All rights reserved.


using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Bodoconsult.Core.Pdf.Helpers
{
    public class ObjectHelper
    {

        public static void MapProperties(object source, object target)
        {

            var sourceType = source.GetType();
            var targetType = target.GetType();

            var propMap = GetMatchingProperties(sourceType, targetType);

            for (var i = 0; i < propMap.Count; i++)

            {

                var prop = propMap[i];

                var sourceValue = prop.SourceProperty.GetValue(source,
                    null);

                prop.TargetProperty.SetValue(target, sourceValue, null);

            }
        }


        public static IList<PropertyMap> GetMatchingProperties(Type sourceType, Type targetType)

        {

            var sourceProperties = sourceType.GetProperties();

            var targetProperties = targetType.GetProperties();



            var properties = (from s in sourceProperties

                from t in targetProperties

                where s.Name == t.Name &&

                      s.CanRead &&

                      t.CanWrite &&
                      s.PropertyType.IsPublic &&
                      t.PropertyType.IsPublic  &&

            s.PropertyType == t.PropertyType &&

            (

                s.PropertyType.IsValueType &&

                t.PropertyType.IsValueType ||

                s.PropertyType == typeof(string) &&

                t.PropertyType == typeof(string)

            )

            select new PropertyMap

            {

                SourceProperty = s,

                TargetProperty = t

            }).ToList();

            return properties;

        }

        public class PropertyMap

        {

            public PropertyInfo SourceProperty { get; set; }

            public PropertyInfo TargetProperty { get; set; }

        }
    }
}
