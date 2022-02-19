namespace SimpleExcelExporter.Annotations
{
  using System;

  public abstract class ResourceBaseAttribute : Attribute
  {
    protected ResourceBaseAttribute(Type resourceType, string resourceName)
    {
      ResourceName = resourceName;
      ResourceType = resourceType;
      Text = ResourceHelper.GetResourceLookup(ResourceType, resourceName);
    }

    public Type ResourceType { get; }

    public string ResourceName { get; }

    public string Text { get; }
  }
}
