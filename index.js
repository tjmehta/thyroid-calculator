var xlsx = require('node-xlsx').default

const worksheets = xlsx.parse(`${__dirname}/input.xlsx`)

let rows = worksheets[0].data
cols = worksheets[0].data[0]

/// utils
function assert(condition, message) {
  if (!condition) throw new Error(message)
}
function get(obj, key) {
  return obj[key]
}
function match(obj, key, regexp) {
  const val = get(obj, key)
  if (val == null) return null
  return regexp.test(val)
}

// calculate data
rows
  .slice(1)
  .map((row, _i) => {
    const i = _i + 1
    // to row object
    return row.reduce((obj, val, i) => {
      const key = cols[i]
      if (key == null) {
        obj.id = val
        return obj
      }
      obj[key] = val
      return obj
    }, {})
  })
  .forEach((obj, i) => {
    /**
     * data transform
     */
    const {
      e_age,
      e_size,
      e_gender,
      e_invasion,
      e_invasion_gross,
      e_invasion_rln,
      e_invasion_post,
      e_multifocal,
      e_surgical,
      e_nodes_6,
      e_nodes_other,
      e_distant,
      e_anaplastic,
      e_tumor,
    } = {
      e_age: get(obj, 'Age at Diagnosis'),
      e_size: get(obj, 'CS Tumor Size'),
      e_gender: match(obj, 'Sex Value', /male/i) ? 'm' : 'f',
      e_invasion: match(obj, 'Lymph-vascular Invasion Value', /1/) ? 'y' : 'n',
      e_invasion_gross: match(obj, 'TNM Path T Value', /3b/i) ? 'y' : 'n',
      e_invasion_rln: match(obj, 'TNM Path T Value', /4a/i) ? 'y' : 'n',
      e_invasion_post: match(obj, 'TNM Path T Value', /4b/i) ? 'y' : 'n',
      e_multifocal: 'n',
      e_surgical: 'y',
      e_nodes_6: match(obj, 'TNM Path N Value', /1a/i) ? 'y' : 'n',
      e_nodes_other: match(obj, 'TNM Path N Value', /1b/i) ? 'y' : 'n',
      e_distant: match(obj, 'TNM Path M Value', /1/i) ? 'y' : 'n',
      e_anaplastic: 'n',
      e_tumor: '',
    }

    let invalid = false
    /**
     * validate required values
     */
    if (e_age == null) {
      console.warn(`Row ${i} Skipped: Age is missing`)
      invalid = true
    }
    if (e_size == null) {
      console.warn(`Row ${i} Skipped: Tumor size is missing`)
      invalid = true
    }
    if (e_invasion == null) {
      console.warn(`Row ${i} Skipped: Invasion (Any) is missing`)
      invalid = true
    }
    if (e_invasion_gross == null) {
      console.warn(`Row ${i} Skipped: Invasion (Gross) is missing`)
      invalid = true
    }
    if (e_invasion_rln == null)
      console.warn(
        `Row ${i} Skipped: Invasion (RLN, Larynx, Trachea, Esophagus) is missing`,
      )
    invalid = true
    if (e_invasion_post == null) {
      console.warn(
        `Row ${i} Skipped: Invasion (Posterior cervical fascia / vessels) is missing`,
      )
      invalid = true
    }
    if (e_multifocal == null) {
      console.warn(`Row ${i} Skipped: Multifocal is missing`)
      invalid = true
    }
    if (e_surgical == null) {
      console.warn(`Row ${i} Skipped: Complete surgical resection is missing`)
      invalid = true
    }
    if (e_nodes_6 == null) {
      console.warn(`Row ${i} Skipped: Nodes (level VI or VII) is missing`)
      invalid = true
    }
    if (e_nodes_other == null) {
      console.warn(`Row ${i} Skipped: Nodes (levels I-V) is missing`)
      invalid = true
    }
    if (e_distant == null) {
      console.warn(`Row ${i} Skipped: Distant metastases is missing`)
      invalid = true
    }
    if (e_anaplastic == null) {
      console.warn(`Row ${i} Skipped: Anaplastic is missing`)
      invalid = true
    }
    // age limit
    if (e_age < 1 || e_age > 120) {
      console.warn(`Row ${i} Skipped: Age should be between 1 and 120 years`)
      invalid = true
    }

    /**
     * validate values
     */
    // tumor grade
    if (e_tumor > 4 || e_tumor < 0) {
      console.warn(`Row ${i} Skipped: Tumor grade should be between 1 and 4`)
      invalid = true
    }
    // tumor limit
    if (e_size < 0.1 || e_size > 50) {
      console.warn(
        `Row ${i} Skipped: Tumor size should be between 0.1 and 50cm`,
      )
      invalid = true
    }

    if (invalid) return

    /**
     * calculate output values
     */
    // TNM
    o_ptnm =
      '' +
      ((e_size <= 1 ? 1 : 0) *
      (e_invasion_gross == 'n' ? 1 : 0) *
      (e_invasion_rln == 'n' ? 1 : 0) *
      (e_invasion_post == 'n' ? 1 : 0)
        ? 'T1a'
        : '') +
      ((e_size > 1 ? 1 : 0) *
      (e_size <= 2 ? 1 : 0) *
      (e_invasion_gross == 'n' ? 1 : 0) *
      (e_invasion_rln == 'n' ? 1 : 0) *
      (e_invasion_post == 'n' ? 1 : 0)
        ? 'T1b'
        : '') +
      ((e_size > 2 ? 1 : 0) *
      (e_size <= 4 ? 1 : 0) *
      (e_invasion_gross == 'n' ? 1 : 0) *
      (e_invasion_rln == 'n' ? 1 : 0) *
      (e_invasion_post == 'n' ? 1 : 0)
        ? 'T2'
        : '') +
      ((e_size > 4 ? 1 : 0) *
      ((e_invasion_gross == 'n' ? 1 : 0) *
        (e_invasion_rln == 'n' ? 1 : 0) *
        (e_invasion_post == 'n' ? 1 : 0))
        ? 'T3a'
        : '') +
      ((e_invasion_gross == 'y' ? 1 : 0) *
      (e_invasion_rln == 'n' ? 1 : 0) *
      (e_invasion_post == 'n' ? 1 : 0)
        ? 'T3b'
        : '') +
      ((e_invasion_rln == 'y' ? 1 : 0) * (e_invasion_post == 'n' ? 1 : 0)
        ? 'T4a'
        : '') +
      (e_invasion_post == 'y' ? 'T4b' : '') +
      (e_multifocal == 'y' ? '(m)' : '(s)') +
      ((e_nodes_6 == 'y' ? 1 : 0) * (e_nodes_other == 'n' ? 1 : 0)
        ? 'N1a'
        : '') +
      (e_nodes_other == 'y' ? 'N1b' : '') +
      ((e_nodes_6 == 'n' ? 1 : 0) * (e_nodes_other == 'n' ? 1 : 0)
        ? 'N0'
        : '') +
      (e_distant == 'y' ? 'M1' : 'M0')

    var e_ptnm = o_ptnm

    // Stage
    o_stage =
      e_anaplastic == 'y'
        ? // Anaplastic is YES
          ((e_ptnm.match(/T1[ab]\(.\)N0M0/) ? 1 : 0) +
          (e_ptnm.match(/T2\(.\)N0M0/) ? 1 : 0) +
          (e_ptnm.match(/T3a\(.\)N0M0/) ? 1 : 0)
            ? 'IVA'
            : '') +
          (((e_distant == 'n' ? 1 : 0) *
          ((e_nodes_6 == 'y' ? 1 : 0) + (e_nodes_other == 'y' ? 1 : 0))
            ? 1
            : 0) +
          ((e_ptnm.match(/T3b\(.\)N0M0/) ? 1 : 0) +
          (e_ptnm.match(/T3b\(.\)N1[ab]M0/) ? 1 : 0) +
          (e_ptnm.match(/T4[ab]\(.\)N0M0/) ? 1 : 0) +
          (e_ptnm.match(/T4[ab]\(.\)N1[ab]M0/) ? 1 : 0)
            ? 1
            : 0)
            ? 'IVB'
            : '') +
          ((e_distant == 'y' ? 1 : 0) ? 'IVC' : '')
        : // Anaplastic is NO
          ((e_age < 55 ? 1 : 0) * (e_distant == 'n' ? 1 : 0) ? 'I' : '') +
          ((e_age < 55 ? 1 : 0) * (e_distant == 'y' ? 1 : 0) ? 'II' : '') +
          ((e_age >= 55 ? 1 : 0) *
          ((e_ptnm.match(/T1[ab]\(.\)N0M0/) ? 1 : 0) +
            (e_ptnm.match(/T2\(.\)N0M0/) ? 1 : 0))
            ? 'I'
            : '') +
          ((e_age >= 55 ? 1 : 0) *
          ((e_ptnm.match(/T1[ab]\(.\)N1[ab]M0/) ? 1 : 0) +
            (e_ptnm.match(/T2\(.\)N1[ab]M0/) ? 1 : 0) +
            (e_ptnm.match(/T3[ab]\(.\)N0M0/) ? 1 : 0) +
            (e_ptnm.match(/T3[ab]\(.\)N1[ab]M0/) ? 1 : 0))
            ? 'II'
            : '') +
          ((e_age >= 55 ? 1 : 0) *
          ((e_ptnm.match(/T4a\(.\)N0M0/) ? 1 : 0) +
            (e_ptnm.match(/T4a\(.\)N1[ab]M0/) ? 1 : 0))
            ? 'III'
            : '') +
          ((e_age >= 55 ? 1 : 0) *
          ((e_ptnm.match(/T4b\(.\)N0M0/) ? 1 : 0) +
            (e_ptnm.match(/T4b\(.\)N1[ab]M0/) ? 1 : 0))
            ? 'IVA'
            : '') +
          ((e_age >= 55 ? 1 : 0) * (e_distant == 'y' ? 1 : 0) ? 'IVB' : '')

    //   MACIS
    var t_macis =
      (e_age < 40 ? 3.1 : 0.08 * e_age) +
      0.3 * e_size +
      (e_invasion == 'n' ? 0 : 1) +
      (e_surgical == 'y' ? 0 : 1) +
      (e_distant == 'y' ? 3 : 0)
    o_macis = Math.round(t_macis * 100) / 100

    //   AGES
    var t_age =
      (e_age < 40 ? 0 : 0.05) * e_age +
      0.2 * e_size +
      (e_invasion == 'n' ? 0 : 1) +
      (e_tumor == 2 ? 1 : 0) +
      (e_tumor == 3 ? 3 : 0) +
      (e_tumor == 4 ? 3 : 0) +
      (e_distant == 'y' ? 3 : 0)
    o_ages = e_tumor.trim() ? Math.round(t_age * 100) / 100 : ''

    // AMES
    o_ames =
      ((e_age < 51 ? 1 : 0) *
      (e_gender == 'f' ? 1 : 0) *
      (e_distant == 'n' ? 1 : 0) *
      (e_invasion == 'n' ? 1 : 0)
        ? 'Low risk'
        : '') +
      ((e_age < 41 ? 1 : 0) *
      (e_gender == 'm' ? 1 : 0) *
      (e_distant == 'n' ? 1 : 0) *
      (e_invasion == 'n' ? 1 : 0)
        ? 'Low risk'
        : '') +
      ((e_distant == 'y' ? 1 : 0) ? 'High risk' : '') +
      ((e_gender == 'f' ? 1 : 0) *
      (e_distant == 'n' ? 1 : 0) *
      (e_invasion == 'y' ? 1 : 0)
        ? 'High risk'
        : '') +
      ((e_gender == 'm' ? 1 : 0) *
      (e_distant == 'n' ? 1 : 0) *
      (e_invasion == 'y' ? 1 : 0)
        ? 'High risk'
        : '') +
      ((e_age >= 51 ? 1 : 0) *
      (e_gender == 'f' ? 1 : 0) *
      (e_distant == 'n' ? 1 : 0) *
      (e_invasion == 'n' ? 1 : 0) *
      (e_size > 5 ? 1 : 0)
        ? 'High risk'
        : '') +
      ((e_age >= 41 ? 1 : 0) *
      (e_gender == 'm' ? 1 : 0) *
      (e_distant == 'n' ? 1 : 0) *
      (e_invasion == 'n' ? 1 : 0) *
      (e_size > 5 ? 1 : 0)
        ? 'High risk'
        : '') +
      ((e_age >= 51 ? 1 : 0) *
      (e_gender == 'f' ? 1 : 0) *
      (e_distant == 'n' ? 1 : 0) *
      (e_invasion == 'n' ? 1 : 0) *
      (e_size <= 5 ? 1 : 0)
        ? 'Low risk'
        : '') +
      ((e_age >= 41 ? 1 : 0) *
      (e_gender == 'm' ? 1 : 0) *
      (e_distant == 'n' ? 1 : 0) *
      (e_invasion == 'n' ? 1 : 0) *
      (e_size <= 5 ? 1 : 0)
        ? 'Low risk'
        : '')

    const calculated = {
      ...obj,
      TNM: o_ptnm,
      Stage: o_stage,
      MACIS: o_macis,
      AGES: o_ages,
      AMES: o_ames,
    }
    console.log('Row', i, 'Success:', calculated)
  })

// console.log(JSON.stringify(rows, null, 2))
