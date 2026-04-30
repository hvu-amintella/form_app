import {
  ARCHIVE_DB_NAME,
  ARCHIVE_FILE_STORE,
} from './constants'

function openArchiveDb() {
  return new Promise((resolve, reject) => {
    if (typeof window === 'undefined' || !window.indexedDB) {
      reject(new Error('IndexedDB indisponible sur cet appareil.'))
      return
    }

    const request = window.indexedDB.open(ARCHIVE_DB_NAME, 1)

    request.onupgradeneeded = () => {
      const db = request.result
      if (!db.objectStoreNames.contains(ARCHIVE_FILE_STORE)) {
        db.createObjectStore(ARCHIVE_FILE_STORE)
      }
    }

    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error || new Error('Impossible d ouvrir la base locale.'))
  })
}

export async function saveArchiveFile(recordId, file) {
  const db = await openArchiveDb()

  await new Promise((resolve, reject) => {
    const transaction = db.transaction(ARCHIVE_FILE_STORE, 'readwrite')
    const store = transaction.objectStore(ARCHIVE_FILE_STORE)
    const request = store.put(file, recordId)

    request.onsuccess = () => resolve()
    request.onerror = () => reject(request.error || new Error('Impossible de sauvegarder le fichier archive.'))
  })

  db.close()
}

export async function getArchiveFile(recordId) {
  const db = await openArchiveDb()

  const file = await new Promise((resolve, reject) => {
    const transaction = db.transaction(ARCHIVE_FILE_STORE, 'readonly')
    const store = transaction.objectStore(ARCHIVE_FILE_STORE)
    const request = store.get(recordId)

    request.onsuccess = () => resolve(request.result || null)
    request.onerror = () => reject(request.error || new Error('Impossible de lire le fichier archive.'))
  })

  db.close()
  return file
}

export async function clearArchiveFiles() {
  const db = await openArchiveDb()

  await new Promise((resolve, reject) => {
    const transaction = db.transaction(ARCHIVE_FILE_STORE, 'readwrite')
    const store = transaction.objectStore(ARCHIVE_FILE_STORE)
    const request = store.clear()

    request.onsuccess = () => resolve()
    request.onerror = () => reject(request.error || new Error('Impossible de vider les fichiers archives.'))
  })

  db.close()
}
