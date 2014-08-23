using System;
using System.Collections.Generic;

namespace ExcelReportsGenerator.Common.DependenyInjection
{
  /// <summary>
  /// The ContainerCatalog interface.
  /// </summary>
  public interface IContainerCatalog
  {
    /// <summary>
    /// Gets the catalog items.
    /// </summary>
    Dictionary<Type, object> CatalogItems { get; }

    /// <summary>
    /// The register.
    /// </summary>
    /// <param name="item">
    /// The item.
    /// </param>
    /// <typeparam name="T">
    /// </typeparam>
    void Register<T>(T item);

    /// <summary>
    /// The unregister.
    /// </summary>
    /// <param name="item">
    /// The item.
    /// </param>
    /// <typeparam name="T">
    /// </typeparam>
    void Unregister<T>(T item);

    /// <summary>
    /// The resolve.
    /// </summary>
    /// <typeparam name="T">
    /// </typeparam>
    /// <returns>
    /// </returns>
    T Resolve<T>();

    /// <summary>
    /// The resolve.
    /// </summary>
    /// <param name="objectType">
    /// The object type.
    /// </param>
    /// <returns>
    /// The <see cref="object"/>.
    /// </returns>
    object Resolve(Type objectType);

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
    T ResolveDependencies<T>(T item);

    /// <summary>
    /// The create instance.
    /// </summary>
    /// <typeparam name="T">
    /// </typeparam>
    /// <returns>
    /// </returns>
    T CreateInstance<T>();

    /// <summary>
    /// The create instance.
    /// </summary>
    /// <param name="objectType">
    /// The object type.
    /// </param>
    /// <returns>
    /// The <see cref="object"/>.
    /// </returns>
    object CreateInstance(Type objectType);
  }
}