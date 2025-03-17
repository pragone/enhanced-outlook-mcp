/**
 * Utility functions for building OData queries for Microsoft Graph API
 */

/**
 * Escape a string value for use in OData filter expressions
 * @param {string} value - The string value to escape
 * @returns {string} - Escaped string
 */
function escapeODataString(value) {
  if (typeof value !== 'string') {
    return value;
  }
  
  // Replace single quotes with two single quotes
  return value.replace(/'/g, "''");
}

/**
 * Format a value for use in OData filter expressions
 * @param {any} value - The value to format
 * @returns {string} - Formatted value
 */
function formatODataValue(value) {
  if (value === null || value === undefined) {
    return 'null';
  }
  
  if (typeof value === 'string') {
    return `'${escapeODataString(value)}'`;
  }
  
  if (value instanceof Date) {
    return value.toISOString();
  }
  
  if (typeof value === 'boolean') {
    return value ? 'true' : 'false';
  }
  
  return String(value);
}

/**
 * Build an OData filter expression
 * @param {Object} criteria - Filter criteria
 * @param {string} operator - Logical operator to join criteria ('and' or 'or')
 * @returns {string} - OData filter expression
 */
function buildFilter(criteria, operator = 'and') {
  if (!criteria || Object.keys(criteria).length === 0) {
    return '';
  }
  
  const filters = [];
  
  for (const [key, value] of Object.entries(criteria)) {
    if (value === undefined) {
      continue;
    }
    
    // Handle complex operators
    if (typeof value === 'object' && value !== null && !Array.isArray(value) && !(value instanceof Date)) {
      for (const [op, opValue] of Object.entries(value)) {
        if (opValue === undefined) {
          continue;
        }
        
        switch (op.toLowerCase()) {
          case 'eq':
            filters.push(`${key} eq ${formatODataValue(opValue)}`);
            break;
          case 'ne':
            filters.push(`${key} ne ${formatODataValue(opValue)}`);
            break;
          case 'gt':
            filters.push(`${key} gt ${formatODataValue(opValue)}`);
            break;
          case 'ge':
            filters.push(`${key} ge ${formatODataValue(opValue)}`);
            break;
          case 'lt':
            filters.push(`${key} lt ${formatODataValue(opValue)}`);
            break;
          case 'le':
            filters.push(`${key} le ${formatODataValue(opValue)}`);
            break;
          case 'contains':
            filters.push(`contains(${key}, ${formatODataValue(opValue)})`);
            break;
          case 'startswith':
            filters.push(`startswith(${key}, ${formatODataValue(opValue)})`);
            break;
          case 'endswith':
            filters.push(`endswith(${key}, ${formatODataValue(opValue)})`);
            break;
          default:
            // Ignore invalid operators
            break;
        }
      }
    } else if (Array.isArray(value)) {
      // Handle array values (IN operator equivalent)
      if (value.length > 0) {
        const formattedValues = value.map(v => formatODataValue(v)).join(',');
        filters.push(`${key} in (${formattedValues})`);
      }
    } else {
      // Default to equality
      filters.push(`${key} eq ${formatODataValue(value)}`);
    }
  }
  
  if (filters.length === 0) {
    return '';
  }
  
  return filters.join(` ${operator} `);
}

/**
 * Build OData $select parameter
 * @param {Array|string} fields - Fields to select
 * @returns {string} - Formatted select parameter
 */
function buildSelect(fields) {
  if (!fields) {
    return '';
  }
  
  if (Array.isArray(fields)) {
    return fields.join(',');
  }
  
  return String(fields);
}

/**
 * Build OData $orderby parameter
 * @param {Array|Object|string} orderBy - Order by specifications
 * @returns {string} - Formatted orderby parameter
 */
function buildOrderBy(orderBy) {
  if (!orderBy) {
    return '';
  }
  
  if (typeof orderBy === 'string') {
    return orderBy;
  }
  
  if (Array.isArray(orderBy)) {
    return orderBy.join(',');
  }
  
  // Handle object format { field: 'asc'/'desc' }
  if (typeof orderBy === 'object') {
    return Object.entries(orderBy)
      .map(([field, direction]) => `${field} ${direction}`)
      .join(',');
  }
  
  return '';
}

/**
 * Build complete OData query parameters
 * @param {Object} options - Query options
 * @returns {Object} - Query parameters for API request
 */
function buildQueryParams(options = {}) {
  const params = {};
  
  // Add filter if provided
  if (options.filter) {
    params['$filter'] = typeof options.filter === 'string' 
      ? options.filter 
      : buildFilter(options.filter, options.filterOperator || 'and');
  }
  
  // Add select if provided
  if (options.select) {
    params['$select'] = buildSelect(options.select);
  }
  
  // Add orderby if provided
  if (options.orderBy) {
    params['$orderby'] = buildOrderBy(options.orderBy);
  }
  
  // Add top if provided
  if (options.top) {
    params['$top'] = options.top;
  }
  
  // Add skip if provided
  if (options.skip) {
    params['$skip'] = options.skip;
  }
  
  // Add expand if provided
  if (options.expand) {
    params['$expand'] = typeof options.expand === 'string'
      ? options.expand
      : options.expand.join(',');
  }
  
  // Add search if provided
  if (options.search) {
    params['$search'] = `"${escapeODataString(options.search)}"`;
  }
  
  // Add count if provided
  if (options.count === true) {
    params['$count'] = 'true';
  }
  
  return params;
}

module.exports = {
  escapeODataString,
  formatODataValue,
  buildFilter,
  buildSelect,
  buildOrderBy,
  buildQueryParams
};