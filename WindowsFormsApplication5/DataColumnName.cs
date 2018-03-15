using System;
using System.Collections.Generic;

namespace WindowsFormsApplication5
{
    [Serializable]
    class DataColumnName : IDisposable
    {
        bool disposed = false;

        public List<string> Chapter_test { get; }
        public List<string> Chapter { get; }
        public List<string> Component { get; }
        public List<string> Graph { get; }

        public DataColumnName(List<string> chapter_test, List<string> chapter, List<string> component, List<string> graph)
        {
            Chapter_test = chapter_test;
            Chapter = chapter;
            Component = component;
            Graph = graph;
        }
        public DataColumnName(List<string> chapter_test, List<string> component, List<string> graph)
        {
            Chapter_test = chapter_test;
            //Chapter = chapter;
            Component = component;
            Graph = graph;
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {

            }
            disposed = true;
        }
    }
}
