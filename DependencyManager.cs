using System;
using System.Collections.Generic;

namespace CalcEngine
{
    public class DependencyComparer : IEqualityComparer<string>
    {
        public bool Equals(string source, string target)
        {
            return string.Equals(source, target);
        }

        public int GetHashCode(string source)
        {
            return source.GetHashCode();
        }
    }

    public class DependencyManager<T>
    {
        #region Private Properties

        /// <summary>
        /// Map of a node and the nodes that depend on it
        /// </summary>
        private Dictionary<T, List<T>> _dependentsMap;

        /// <summary>
        /// Map of a node and the number of nodes that point to it
        /// </summary>
        private Dictionary<T, int> _precedentsMap;

        private IEqualityComparer<T> _equalityComparer;

        #endregion Private Properties

        #region Public Properties

        public string Precedents
        {
            get
            {
                List<string> list = new List<string>();
                foreach (KeyValuePair<T, int> pair in _precedentsMap)
                {
                    list.Add(pair.ToString());
                }
                return string.Join(Environment.NewLine, list);
            }
        }

        public string DependencyGraph
        {
            get
            {
                List<string> lines = new List<string>();
                foreach (KeyValuePair<T, List<T>> pair in _dependentsMap)
                {
                    T propertyName = pair.Key;
                    string dependents = FormatValues(pair.Value);
                    lines.Add(string.Format("* {0} -> {1}", propertyName, dependents));
                }
                return string.Join(Environment.NewLine, lines);
            }
        }

        public int DependencyCount
        {
            get { return _dependentsMap.Count; }
        }

        #endregion Public Properties

        #region Constructor

        public DependencyManager(IEqualityComparer<T> comparer)
        {
            _equalityComparer = comparer;
            _dependentsMap = new Dictionary<T, List<T>>(_equalityComparer);
            _precedentsMap = new Dictionary<T, int>(_equalityComparer);
        }

        #endregion Constructor

        #region Public Methods

        /// <summary>
        /// Create a dependency list with only the dependents of the given targets
        /// </summary>
        /// <param name="tails">The targets.</param>
        /// <returns></returns>
        public DependencyManager<T> CloneDependents(T[] targets)
        {
            List<T> seenNodes = new List<T>();
            DependencyManager<T> copy = new DependencyManager<T>(_equalityComparer);

            foreach (T target in targets)
            {
                CloneDependentsInternal(target, copy, seenNodes);
            }

            return copy;
        }

        public T[] GetTargets()
        {
            T[] arr = new T[_dependentsMap.Keys.Count];

            _dependentsMap.Keys.CopyTo(arr, 0);
            return arr;
        }

        /// <summary>
        /// Clears both the dependents and precedents maps.
        /// </summary>
        public void Clear()
        {
            _dependentsMap.Clear();
            _precedentsMap.Clear();
        }

        public void ReplaceDependency(T old, T replaceWith)
        {
            List<T> value = _dependentsMap[old];

            _dependentsMap.Remove(old);
            _dependentsMap.Add(replaceWith, value);

            foreach (List<T> dependents in _dependentsMap.Values)
            {
                if (dependents.Contains(old) == true)
                {
                    dependents.Remove(old);
                    dependents.Add(replaceWith);
                }
            }
        }

        public void AddTarget(T target)
        {
            if (_dependentsMap.ContainsKey(target) == false)
            {
                _dependentsMap.Add(target, new List<T>());
            }
        }

        public void AddDepedency(T target, T dependent)
        {
            List<T> dependents = GetTargetDependents(target);
            //ETM_VERIFY: If ignoring direct dependents, meaning target property in expression, causing any problems
            if (dependents.Contains(dependent) == false && !_equalityComparer.Equals(target, dependent))
            {
                List<T> dependentDependents = new List<T>();
                GetDependentsRecursive(dependent, dependentDependents);
                if (dependentDependents.Contains(target) == true)
                    throw new CircularReferenceException(string.Format("adding dependent '{0}' to target '{1}'", dependent.ToString(), target.ToString()));
                dependents.Add(dependent);
                AddPrecedent(dependent);
            }
        }

        public void RemoveDependency(T target, T dependent)
        {
            IList<T> dependents = GetTargetDependents(target);
            RemoveHead(dependent, dependents);
        }

        public void Remove(T[] targets)
        {
            foreach (List<T> dependents in _dependentsMap.Values)
            {
                foreach (T target in targets)
                {
                    RemoveHead(target, dependents);
                }
            }

            foreach (T target in targets)
            {
                _dependentsMap.Remove(target);
            }
        }

        public List<T> GetDirectDependents(T target)
        {
            return GetTargetDependents(target);
        }

        public T[] GetDependents(T target)
        {
            List<T> dependents = new List<T>();
            GetDependentsRecursive(target, dependents);
            T[] arr = new T[dependents.Count];
            dependents.CopyTo(arr, 0);
            return arr;
        }

        public void GetDirectPrecedents(T dependent, IList<T> dest)
        {
            foreach (T target in _dependentsMap.Keys)
            {
                IList<T> dependents = GetTargetDependents(target);
                if (dependents.Contains(dependent) == true)
                {
                    dest.Add(target);
                }
            }
        }

        public bool HasPrecedents(T dependent)
        {
            return _precedentsMap.ContainsKey(dependent);
        }

        public bool HasDependents(T target)
        {
            IList<T> dependents = GetTargetDependents(target);
            return dependents.Count > 0;
        }

        /// <summary>
        /// Add all nodes that don't have any incoming edges into a queue
        /// </summary>
        /// <param name="rootTargets">The root tails.</param>
        /// <returns></returns>
        public Queue<T> GetSources(T[] rootTargets)
        {
            Queue<T> q = new Queue<T>();

            foreach (T rootTarget in rootTargets)
            {
                if (HasPrecedents(rootTarget) == false)
                {
                    q.Enqueue(rootTarget);
                }
            }

            return q;
        }

        public string TopologicalSort
        {
            get
            {
                List<string> lines = new List<string>();
                foreach (KeyValuePair<T, List<T>> pair in _dependentsMap)
                {
                    T propertyName = pair.Key;
                    string dependents = FormatValues(pair.Value);
                    lines.Add(string.Format("{0} -> {1}", propertyName, dependents));
                }
                return string.Join(Environment.NewLine, lines);
            }
        }

        #endregion Public Methods

        #region Private Methods

        private List<T> GetTargetDependents(T target)
        {
            List<T> value = null;
            if (target != null)
                _dependentsMap.TryGetValue(target, out value);
            else
            {
                value = new List<T>();
                foreach (T key in _dependentsMap.Keys)
                    value.Add(key);
            }

            return value;
        }

        private void CloneDependentsInternal(T target, DependencyManager<T> dpManager, IList<T> seenNodes)
        {
            if (seenNodes.Contains(target) == true)
            {
                // We've already added this node so just return
                return;
            }
            else
            {
                // Haven't seen this node yet; mark it as visited
                seenNodes.Add(target);
                dpManager.AddTarget(target);
            }

            // Do the recursive add
            IList<T> dependents = GetTargetDependents(target);
            if (dependents != null)
            {
                foreach (T dependent in dependents)
                {
                    dpManager.AddDepedency(target, dependent);
                    CloneDependentsInternal(dependent, dpManager, seenNodes);
                }
            }
        }

        private void RemoveHead(T dependent, IList<T> dependents)
        {
            if (dependents.Remove(dependent) == true)
            {
                RemovePrecedent(dependent);
            }
        }

        private void GetDependentsRecursive(T target, List<T> dependents)
        {
            if (dependents.Contains(target) == false)
                dependents.Add(target);

            List<T> directDependents = GetTargetDependents(target);
            if (directDependents != null)
            {
                foreach (T newTarget in directDependents)
                {
                    GetDependentsRecursive(newTarget, dependents);
                }
            }
        }

        private void AddPrecedent(T dependent)
        {
            int count = 0;
            _precedentsMap.TryGetValue(dependent, out count);
            _precedentsMap[dependent] = count + 1;
        }

        private void RemovePrecedent(T dependent)
        {
            int count = _precedentsMap[dependent] - 1;
            if (count == 0)
                _precedentsMap.Remove(dependent);
            else
                _precedentsMap[dependent] = count;
        }

        private string FormatValues(ICollection<T> values)
        {
            string[] strings = new string[values.Count];
            T[] keys = new T[values.Count];
            values.CopyTo(keys, 0);

            for (int i = 0; i <= keys.Length - 1; i++)
            {
                strings[i] = keys[i].ToString();
            }

            return (strings.Length == 0) ? "<empty>" : string.Join(",", strings);
        }

        #endregion Private Methods
    }
}