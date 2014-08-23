using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportsGenerator.Common.DependenyInjection
{
  /// <summary>
  /// The container catalog.
  /// </summary>
  public class ContainerCatalog : IContainerCatalog
  {
    /// <summary>
    /// The _catalog items.
    /// </summary>
    private readonly Dictionary<Type, object> _catalogItems = new Dictionary<Type, object>();

    #region IContainerCatalog Members

    /// <summary>
    /// Gets the catalog items.
    /// </summary>
    public Dictionary<Type, object> CatalogItems
    {
      get { return this._catalogItems; }
    }

    /// <summary>
    /// The register.
    /// </summary>
    /// <param name="item">
    /// The item.
    /// </param>
    /// <typeparam name="T">
    /// </typeparam>
    /// <exception cref="Exception">
    /// </exception>
    public void Register<T>(T item)
    {
      if (this._catalogItems.ContainsKey(typeof(T)))
      {
        throw new Exception("Item already exists in catalogue.");
      }

      this._catalogItems.Add(typeof(T), item);
    }

    /// <summary>
    /// The unregister.
    /// </summary>
    /// <param name="item">
    /// The item.
    /// </param>
    /// <typeparam name="T">
    /// </typeparam>
    public void Unregister<T>(T item)
    {
      if (this._catalogItems.ContainsKey(typeof(T)))
      {
        this._catalogItems.Remove(typeof(T));
      }
    }

    /// <summary>
    /// Unregisters all.
    /// </summary>
    public void UnregisterAll()
    {
      if (this._catalogItems != null)
      {
        this._catalogItems.Clear();
      }
    }

    /// <summary>
    /// The resolve.
    /// </summary>
    /// <typeparam name="T">
    /// </typeparam>
    /// <returns>
    /// </returns>
    public T Resolve<T>()
    {
      if (this._catalogItems.ContainsKey(typeof(T)))
      {
        return (T)this._catalogItems[typeof(T)];
      }

      return default(T);
    }

    /// <summary>
    /// The resolve.
    /// </summary>
    /// <param name="objectType">
    /// The object type.
    /// </param>
    /// <returns>
    /// The <see cref="object"/>.
    /// </returns>
    public object Resolve(Type objectType)
    {
      if (this._catalogItems.ContainsKey(objectType))
      {
        return this._catalogItems[objectType];
      }

      return null;
    }

    /// <summary>
    /// The resolve dependencies.
    /// </summary>
    /// <param name="item">
    /// The item.
    /// </param>
    /// <typeparam name="T">
    /// </typeparam>
    /// <returns>
    /// </returns>
    public T ResolveDependencies<T>(T item)
    {
      if (item != null)
      {
        IEnumerable<PropertyInfo> importProperties = from propertyItem in typeof(T).GetProperties()
                                                     from attributeItem in propertyItem.GetCustomAttributes(false)
                                                     where attributeItem is ImportAttribute
                                                     select propertyItem;

        foreach (PropertyInfo importProperty in importProperties)
        {
          importProperty.SetValue(item, this.Resolve(importProperty.PropertyType), new object[] { });
        }

        return item;
      }

      return default(T);
    }

    /// <summary>
    /// The create instance.
    /// </summary>
    /// <typeparam name="T">
    /// </typeparam>
    /// <returns>
    /// </returns>
    public T CreateInstance<T>()
    {
      return (T)this.CreateInstance(typeof(T));
    }

    /// <summary>
    /// The create instance.
    /// </summary>
    /// <param name="objectType">
    /// The object type.
    /// </param>
    /// <returns>
    /// The <see cref="object"/>.
    /// </returns>
    public object CreateInstance(Type objectType)
    {
      object instance = Activator.CreateInstance(objectType);

      return this.ResolveDependencies(instance);
    }

    #endregion

    /// <summary>
    /// The resolve dependencies.
    /// </summary>
    /// <param name="item">
    /// The item.
    /// </param>
    /// <returns>
    /// The <see cref="object"/>.
    /// </returns>
    public object ResolveDependencies(object item)
    {
      if (item != null)
      {
        IEnumerable<PropertyInfo> importProperties = from propertyItem in item.GetType().GetProperties()
                                                     from attributeItem in propertyItem.GetCustomAttributes(false)
                                                     where attributeItem is ImportAttribute
                                                     select propertyItem;

        foreach (PropertyInfo importProperty in importProperties)
        {
          importProperty.SetValue(item, this.Resolve(importProperty.PropertyType), new object[] { });
        }

        return item;
      }

      return null;
    }
  }
}
