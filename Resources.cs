using System;
using System.Reflection;

internal class Resources
{
	internal delegate void Delegate4(object o);

	internal static Module module_0;

	internal static void p3dZ3HyyElhA7(int typemdt)
	{
		Type type = Resources.module_0.ResolveType(33554432 + typemdt);
		FieldInfo[] fields = type.GetFields();
		for (int i = 0; i < fields.Length; i++)
		{
			FieldInfo fieldInfo = fields[i];
			MethodInfo method = (MethodInfo)Resources.module_0.ResolveMethod(fieldInfo.MetadataToken + 100663296);
			fieldInfo.SetValue(null, (MulticastDelegate)Delegate.CreateDelegate(type, method));
		}
	}

}
