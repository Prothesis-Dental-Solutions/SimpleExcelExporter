namespace SimpleExcelExporter
{
  using System;
  using System.Reflection;

  public static class ResourceHelper
  {
    public static string GetResourceLookup(Type resourceType, string resourceName)
    {
      var property = resourceType.GetProperty(resourceName, BindingFlags.Public | BindingFlags.Static);
      if (property == null)
      {
        throw new InvalidOperationException($"The resource type [{resourceType}] does not have a property named {resourceName}");
      }

      if (property.PropertyType != typeof(string))
      {
        throw new InvalidOperationException("Resource Property is Not String Type");
      }

      return (string?)property.GetValue(null, null) ?? string.Empty;
    }
  }
}
