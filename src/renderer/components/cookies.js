
const TokenKey = 'EMAILER'

export function getCookie(guid) {
  return localStorage[`${TokenKey}-${guid}`];
}

export function setCookie(guid,token) {
  localStorage[`${TokenKey}-${guid}`] = token;
}

